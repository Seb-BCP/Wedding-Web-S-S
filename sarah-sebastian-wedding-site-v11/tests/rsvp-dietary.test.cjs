const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const project = process.env.WEDDING_TEST_PROJECT || path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(project, 'index.html'), 'utf8');
const pageScript = html.match(/<script>([\s\S]*?)<\/script>/)[1];
const rsvpScript = pageScript.slice(pageScript.indexOf('    function setStatus('), pageScript.lastIndexOf('    startEnvelopeReveal();'));
const backendScript = fs.readFileSync(path.join(project, 'google-apps-script.gs'), 'utf8');
const apiUrl = html.match(/const RSVP_API_URL = "([^"]+)"/)[1];

// The production doGet/doPost functions run unchanged against in-memory sheets.
// These tests never access Google or submit real guest responses.
function makeBackend() {
  const guests = [
    ['Invite ID', 'Primary Guest', 'Guest 2', 'Guest 3', 'Email', 'Magic Code', 'RSVP Link'],
    ['TEST001', 'Alex Test', 'Jamie Test', 'Casey Test', 'rsvp-test@example.invalid', 'test-code', ''],
  ];
  const rows = [
    ['Timestamp', 'Invite ID', 'Guest Name', 'Attending', 'Dietary Requirements', 'Email', 'Notes'],
    ['earlier', 'OTHER001', 'Other Guest', 'Yes', 'Vegan', 'other@example.invalid', ''],
  ];
  const accessedIds = [];
  function sheet(data) {
    return {
      getDataRange: () => ({ getValues: () => data.map(row => [...row]) }),
      getLastRow: () => data.length,
      getLastColumn: () => data[0].length,
      deleteRow: row => data.splice(row - 1, 1),
      getRange(row, column, height, width) {
        return {
          getValues: () => data.slice(row - 1, row - 1 + height).map(values => values.slice(column - 1, column - 1 + width)),
          setValues(values) {
            assert.equal(values.length, height);
            values.forEach((value, index) => {
              assert.equal(value.length, width);
              data[row - 1 + index] ||= [];
              data[row - 1 + index].splice(column - 1, width, ...value);
            });
          },
        };
      },
    };
  }
  const sheets = { Guests: sheet(guests), RSVPs: sheet(rows) };
  const context = vm.createContext({
    console,
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    SpreadsheetApp: {
      openById(id) { accessedIds.push(id); return { getSheetByName: name => sheets[name] }; },
    },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput: text => ({ text, setMimeType() { return this; } }),
    },
  });
  vm.runInContext(backendScript, context);
  return {
    rows, accessedIds,
    get: () => JSON.parse(context.doGet({ parameter: { invite: 'test-code' } }).text),
    post: body => JSON.parse(context.doPost({ postData: { contents: body } }).text),
  };
}

function makeForm(backend, options = {}) {
  const registry = new Map();
  function element(tag) {
    const el = {
      tag, children: [], events: {}, value: '', checked: false, hidden: false,
      disabled: false, required: false, validationMessage: '',
      append(...children) { children.forEach(child => { child.parent = this; this.children.push(child); }); },
      replaceChildren(...children) { this.children = []; this.append(...children); },
      addEventListener(event, handler) { this.events[event] = handler; },
      setAttribute(name, value) { this[name] = value; },
      setCustomValidity(message) { this.validationMessage = message; },
    };
    el.classList = { add(name) { el.className = ((el.className || '') + ' ' + name).trim(); } };
    Object.defineProperty(el, 'id', { set(id) { this._id = id; registry.set(id, this); }, get() { return this._id; } });
    return el;
  }
  const form = element('form');
  const submitButton = element('button');
  const guestResponsesEl = element('div');
  const notesEl = element('textarea');
  const statusEl = element('div');
  form.elements = { website: { value: '' } };
  form.append(guestResponsesEl, notesEl, submitButton);
  form.reportValidity = () => [...registry.values()].every(el => {
    if (el.disabled) return true;
    if (el.validationMessage) return false;
    if (!el.required) return true;
    if (el.type === 'radio') return [...registry.values()].some(other => other.name === el.name && other.checked);
    return el.value !== '';
  });
  const posts = [];
  const context = vm.createContext({
    form, submitButton, guestResponsesEl, notesEl, statusEl,
    invitationSection: element('section'), invitationMessageEl: element('p'),
    inviteCode: 'test-code', RSVP_API_URL: apiUrl,
    document: { createElement: element, getElementById: id => registry.get(id) },
    fetch: async (url, request) => {
      assert.equal(url, apiUrl);
      assert.equal(request.method, 'POST');
      assert.equal(request.headers['Content-Type'], 'text/plain;charset=utf-8');
      posts.push(JSON.parse(request.body));
      return { ok: true, json: async () => options.fail ? { ok: false } : backend.post(request.body) };
    },
  });
  vm.runInContext('let invitation = null; let submitting = false;\n' + rsvpScript, context);
  const lookup = backend.get();
  assert.equal(lookup.valid, true);
  context.renderInvitation(lookup);
  return {
    form, posts, notesEl, statusEl, submitButton,
    field: index => registry.get('dietary-' + index),
    setDietary(index, value) {
      const field = registry.get('dietary-' + index);
      field.value = value;
      field.events.input();
    },
    attend(index, yes) {
      const input = registry.get('attending-' + index + '-yes');
      input.checked = yes;
      registry.get('attending-' + index + '-no').checked = !yes;
      input.parent.parent.parent.events.change();
    },
    submit: () => form.events.submit({ preventDefault() {} }),
  };
}

test('dietary question is visible and an attending guest must answer, including after whitespace input', async () => {
  const backend = makeBackend();
  const form = makeForm(backend);
  assert.equal(form.field(0).parent.hidden, false);
  assert.equal(form.field(0).tag, 'textarea');
  form.attend(0, true);
  form.attend(1, false);
  form.attend(2, false);
  assert.equal(form.field(0).required, true);
  await form.submit();
  assert.equal(form.posts.length, 0);
  form.setDietary(0, '   ');
  await form.submit();
  assert.equal(form.posts.length, 0);
  assert.match(form.field(0).validationMessage, /None/);
  form.setDietary(0, 'None');
  await form.submit();
  assert.equal(form.posts.length, 1);
  assert.equal(backend.rows[2][4], 'None');
});

test('actual form payload writes the right dietary answer to column E for each guest', async () => {
  const backend = makeBackend();
  const form = makeForm(backend);
  form.attend(0, true);
  form.setDietary(0, '  Vegetarian\nNo nuts  ');
  form.setDietary(1, 'Gluten-free');
  form.attend(1, false);
  form.attend(2, true);
  form.setDietary(2, 'None');
  form.notesEl.value = 'Looking forward to celebrating.';
  await form.submit();
  assert.deepEqual(form.posts[0].responses, [
    { name: 'Alex Test', attending: true, dietary: 'Vegetarian\nNo nuts' },
    { name: 'Jamie Test', attending: false, dietary: '' },
    { name: 'Casey Test', attending: true, dietary: 'None' },
  ]);
  assert.equal(form.posts[0].invite, 'test-code');
  assert.equal(form.posts[0].email, undefined, 'Email is supplied by the invitation backend');
  assert.deepEqual(backend.rows.slice(2).map(row => row.slice(1)), [
    ['TEST001', 'Alex Test', 'Yes', 'Vegetarian\nNo nuts', 'rsvp-test@example.invalid', 'Looking forward to celebrating.'],
    ['TEST001', 'Jamie Test', 'No', '', 'rsvp-test@example.invalid', 'Looking forward to celebrating.'],
    ['TEST001', 'Casey Test', 'Yes', 'None', 'rsvp-test@example.invalid', 'Looking forward to celebrating.'],
  ]);
  assert(backend.accessedIds.every(id => id === '1XsNAFnWTsdNTvb89ndlO3-oWrvR6WTzRxpt--qnVa6Y'));
  assert.equal(form.form.hidden, true);
  assert.match(form.statusEl.textContent, /RSVP has been sent/);
});

test('resubmitting updates dietary answers without duplicating guests or altering other invitations', async () => {
  const backend = makeBackend();
  const first = makeForm(backend);
  first.attend(0, true); first.setDietary(0, 'Vegetarian');
  first.attend(1, false); first.attend(2, false);
  await first.submit();
  const second = makeForm(backend);
  second.attend(0, true); second.setDietary(0, 'Vegan');
  second.attend(1, true); second.setDietary(1, 'None');
  second.attend(2, false);
  await second.submit();
  assert.equal(backend.rows.length, 5);
  assert.equal(backend.rows[1][1], 'OTHER001');
  assert.equal(backend.rows[1][4], 'Vegan');
  assert.deepEqual(backend.rows.slice(2).map(row => row[4]), ['Vegan', 'None', '']);
});

test('declining all guests needs no dietary answers and clears previous input', async () => {
  const backend = makeBackend();
  const form = makeForm(backend);
  for (let index = 0; index < 3; index++) {
    form.setDietary(index, 'No dairy');
    form.attend(index, false);
    assert.equal(form.field(index).disabled, true);
    assert.equal(form.field(index).required, false);
    assert.equal(form.field(index).parent.hidden, true);
  }
  await form.submit();
  assert.equal(form.posts.length, 1);
  assert(backend.rows.slice(2).every(row => row[3] === 'No' && row[4] === ''));
});

test('backend rejection preserves the form and does not show a successful RSVP', async () => {
  const backend = makeBackend();
  const form = makeForm(backend, { fail: true });
  form.attend(0, true); form.setDietary(0, 'None');
  form.attend(1, false); form.attend(2, false);
  await form.submit();
  assert.equal(form.form.hidden, false);
  assert.match(form.statusEl.textContent, /Error sending RSVP/);
  assert.equal(form.submitButton.disabled, false);
  assert.equal(backend.rows.length, 2);
});
