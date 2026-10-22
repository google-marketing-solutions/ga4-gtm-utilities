/**
 * @fileoverview Mock of the Apps Script `Ui` class (SpreadsheetApp.getUi())
 * plus a host-side controller that records alerts, prompts, toasts, menus and
 * dialogs and lets tests script the user's responses.
 */
'use strict';

const BUTTON = Object.freeze({OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO', CLOSE: 'CLOSE'});
const BUTTON_SET = Object.freeze({
  OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO', YES_NO_CANCEL: 'YES_NO_CANCEL',
});

/** Default button returned for each button set when no response is queued. */
const DEFAULT_RESPONSE = Object.freeze({
  OK: BUTTON.OK, OK_CANCEL: BUTTON.OK, YES_NO: BUTTON.YES, YES_NO_CANCEL: BUTTON.YES,
});

/** Buttons that are valid answers for each button set (CLOSE is always valid). */
const VALID_BUTTONS = Object.freeze({
  OK: ['OK', 'CLOSE'],
  OK_CANCEL: ['OK', 'CANCEL', 'CLOSE'],
  YES_NO: ['YES', 'NO', 'CLOSE'],
  YES_NO_CANCEL: ['YES', 'NO', 'CANCEL', 'CLOSE'],
});

/**
 * Defines non-enumerable properties (keeps mocks clean for JSON/console).
 * @param {!Object} obj
 * @param {!Object} props
 */
function defineHidden(obj, props) {
  for (const [key, value] of Object.entries(props)) {
    Object.defineProperty(obj, key, {value, writable: true, configurable: true, enumerable: false});
  }
}

/**
 * Menu builder mock (Ui.createMenu / Ui.createAddonMenu).
 */
class MenuMock {
  /**
   * @param {!UiController} controller
   * @param {string} caption
   * @param {boolean} isAddon
   */
  constructor(controller, caption, isAddon) {
    defineHidden(this, {_controller: controller, _caption: caption, _isAddon: isAddon, _items: []});
  }

  /**
   * @param {string} caption
   * @param {string} functionName
   * @return {!MenuMock}
   */
  addItem(caption, functionName) {
    if (typeof functionName !== 'string') {
      throw this._controller.realm.error(
          `The parameters (${typeof caption},${typeof functionName}) don't match the method ` +
          'signature for Ui.Menu.addItem.');
    }
    this._items.push({type: 'item', caption: String(caption), functionName});
    return this;
  }

  /** @return {!MenuMock} */
  addSeparator() {
    this._items.push({type: 'separator'});
    return this;
  }

  /**
   * @param {!MenuMock} menu
   * @return {!MenuMock}
   */
  addSubMenu(menu) {
    if (!(menu instanceof MenuMock)) {
      throw this._controller.realm.error(
          'The parameters (' + typeof menu + ') don\'t match the method signature for Ui.Menu.addSubMenu.');
    }
    this._items.push({type: 'submenu', caption: menu._caption, menu});
    return this;
  }

  /** Adds the menu to the (recorded) spreadsheet UI. */
  addToUi() {
    const tree = this._toTree();
    const existing = this._controller.menus.findIndex((m) => m.caption === tree.caption);
    if (existing >= 0) this._controller.menus.splice(existing, 1, tree);
    else this._controller.menus.push(tree);
  }

  /** @return {string} */
  toString() {
    return 'Menu';
  }

  /**
   * Serializes the menu (recursively) into plain host objects.
   * @return {!Object}
   */
  _toTree() {
    return {
      caption: this._caption,
      addon: this._isAddon,
      items: this._items.map((item) => {
        if (item.type === 'submenu') return {type: 'submenu', ...item.menu._toTree()};
        return {...item};
      }),
    };
  }
}

/**
 * Prompt response mock.
 */
class PromptResponseMock {
  /**
   * @param {string} text
   * @param {string} button
   */
  constructor(text, button) {
    defineHidden(this, {_text: text, _button: button});
  }

  /** @return {string} */
  getResponseText() {
    return this._text;
  }

  /** @return {string} */
  getSelectedButton() {
    return this._button;
  }

  /** @return {string} */
  toString() {
    return 'PromptResponse';
  }
}

/**
 * Minimal HtmlOutput-like recording for dialogs and sidebars.
 * @param {*} html
 * @return {!Object}
 */
function describeHtml(html) {
  if (html && typeof html.getContent === 'function') {
    return {
      content: html.getContent(),
      title: typeof html.getTitle === 'function' ? html.getTitle() : undefined,
      width: typeof html.getWidth === 'function' ? html.getWidth() : undefined,
      height: typeof html.getHeight === 'function' ? html.getHeight() : undefined,
    };
  }
  return {content: String(html)};
}

/**
 * Host-side controller: owns the recordings and the scripted response queue.
 * Exposed to tests as `runtime.ui`.
 */
class UiController {
  /**
   * @param {!Object} realm Realm helpers (see realm.js).
   * @param {!Object=} options
   * @param {!Array=} options.responses Initial response queue.
   * @param {function(!Object): *=} options.handler Fallback responder called
   *     with the dialog record when the queue is empty.
   */
  constructor(realm, options = {}) {
    this.realm = realm;
    /** @type {!Array<!Object>} All alert() calls, in order. */
    this.alerts = [];
    /** @type {!Array<!Object>} All prompt() calls, in order. */
    this.prompts = [];
    /** @type {!Array<!Object>} All Spreadsheet.toast() calls, in order. */
    this.toasts = [];
    /** @type {!Array<!Object>} Menus added with addToUi(). */
    this.menus = [];
    /** @type {!Array<!Object>} showSidebar/showModalDialog/showModelessDialog calls. */
    this.dialogs = [];
    /** @type {!Array<!Object>} Every interaction (alerts+prompts+dialogs) in order. */
    this.history = [];
    /** @type {!Array<*>} Scripted responses consumed by alert()/prompt(). */
    this.responses = [...(options.responses || [])];
    this.handler = options.handler || null;
    this.instance = new UiMock(this);
  }

  /**
   * Queues scripted responses. Each entry is one of:
   *  - a Button name ('OK', 'YES', 'NO', 'CANCEL', 'CLOSE') for alert();
   *  - a string for prompt() (the typed text, button OK);
   *  - {text, button} for prompt(), {button} for alert();
   *  - a function(dialogRecord) returning any of the above.
   * @param {...*} responses
   * @return {!UiController}
   */
  queueResponse(...responses) {
    this.responses.push(...responses);
    return this;
  }

  /** Clears recordings and the response queue. */
  reset() {
    this.alerts.length = 0;
    this.prompts.length = 0;
    this.toasts.length = 0;
    this.menus.length = 0;
    this.dialogs.length = 0;
    this.history.length = 0;
    this.responses.length = 0;
  }

  /**
   * Flattens all recorded menus into a list of {path, caption, functionName}.
   * @return {!Array<{path: string, caption: string, functionName: string}>}
   */
  menuItems() {
    const out = [];
    const walk = (items, prefix) => {
      for (const item of items) {
        if (item.type === 'item') {
          out.push({path: `${prefix} > ${item.caption}`, caption: item.caption, functionName: item.functionName});
        } else if (item.type === 'submenu') {
          walk(item.items, `${prefix} > ${item.caption}`);
        }
      }
    };
    for (const menu of this.menus) walk(menu.items, menu.caption);
    return out;
  }

  /** @return {!Array<string>} Function names referenced by menu items. */
  menuFunctionNames() {
    return [...new Set(this.menuItems().map((i) => i.functionName))];
  }

  /**
   * Finds a menu item by caption (exact) or path suffix.
   * @param {string} captionOrPath
   * @return {?{path: string, caption: string, functionName: string}}
   */
  findMenuItem(captionOrPath) {
    return this.menuItems().find(
        (i) => i.caption === captionOrPath || i.path === captionOrPath ||
        i.path.endsWith(` > ${captionOrPath}`)) || null;
  }

  /**
   * Pulls the next scripted response for a dialog.
   * @param {!Object} record
   * @return {*} undefined when nothing is queued.
   */
  _nextResponse(record) {
    let response;
    if (this.responses.length > 0) response = this.responses.shift();
    else if (this.handler) response = this.handler(record);
    if (typeof response === 'function') response = response(record);
    return response;
  }
}

/**
 * The `Ui` object returned by SpreadsheetApp.getUi().
 */
class UiMock {
  /** @param {!UiController} controller */
  constructor(controller) {
    defineHidden(this, {_controller: controller});
    this.Button = BUTTON;
    this.ButtonSet = BUTTON_SET;
  }

  /**
   * alert(prompt) | alert(prompt, buttons) | alert(title, prompt, buttons)
   * @param {...*} args
   * @return {string} The Button the (scripted) user clicked.
   */
  alert(...args) {
    let title = null;
    let prompt;
    let buttons = BUTTON_SET.OK;
    if (args.length >= 3) [title, prompt, buttons] = args;
    else if (args.length === 2) [prompt, buttons] = args;
    else [prompt] = args;
    if (args.length === 0) {
      throw this._controller.realm.error(
          'The parameters () don\'t match the method signature for Ui.alert.');
    }
    if (!VALID_BUTTONS[buttons]) {
      throw this._controller.realm.error(
          `The parameters (String,${typeof buttons}) don't match the method signature for Ui.alert.`);
    }
    const record = {type: 'alert', title, prompt: String(prompt), buttons, response: null};
    let response = this._controller._nextResponse(record);
    if (response && typeof response === 'object') response = response.button;
    if (response === undefined || response === null) response = DEFAULT_RESPONSE[buttons];
    if (!VALID_BUTTONS[buttons].includes(response)) {
      throw new Error(`UiMock: scripted response "${response}" is not valid for ButtonSet.${buttons} ` +
          `(alert: ${JSON.stringify(record.prompt).slice(0, 120)})`);
    }
    record.response = response;
    this._controller.alerts.push(record);
    this._controller.history.push(record);
    return response;
  }

  /**
   * prompt(prompt) | prompt(prompt, buttons) | prompt(title, prompt, buttons)
   * @param {...*} args
   * @return {!PromptResponseMock}
   */
  prompt(...args) {
    let title = null;
    let prompt;
    let buttons = BUTTON_SET.OK;
    if (args.length >= 3) [title, prompt, buttons] = args;
    else if (args.length === 2) [prompt, buttons] = args;
    else [prompt] = args;
    if (!VALID_BUTTONS[buttons]) {
      throw this._controller.realm.error(
          `The parameters (String,${typeof buttons}) don't match the method signature for Ui.prompt.`);
    }
    const record = {type: 'prompt', title, prompt: String(prompt), buttons, response: null, text: null};
    const raw = this._controller._nextResponse(record);
    let text = '';
    let button = DEFAULT_RESPONSE[buttons];
    if (typeof raw === 'string') {
      text = raw;
    } else if (raw && typeof raw === 'object') {
      if (raw.text !== undefined) text = String(raw.text);
      if (raw.button !== undefined) button = raw.button;
    }
    if (!VALID_BUTTONS[buttons].includes(button)) {
      throw new Error(`UiMock: scripted button "${button}" is not valid for ButtonSet.${buttons} ` +
          `(prompt: ${JSON.stringify(record.prompt).slice(0, 120)})`);
    }
    record.response = button;
    record.text = text;
    this._controller.prompts.push(record);
    this._controller.history.push(record);
    return new PromptResponseMock(text, button);
  }

  /**
   * @param {string} caption
   * @return {!MenuMock}
   */
  createMenu(caption) {
    return new MenuMock(this._controller, String(caption), false);
  }

  /** @return {!MenuMock} */
  createAddonMenu() {
    return new MenuMock(this._controller, 'Add-on', true);
  }

  /** @param {*} html */
  showSidebar(html) {
    this._recordDialog('sidebar', html);
  }

  /**
   * @param {*} html
   * @param {string} title
   */
  showModalDialog(html, title) {
    this._recordDialog('modal', html, title);
  }

  /**
   * @param {*} html
   * @param {string} title
   */
  showModelessDialog(html, title) {
    this._recordDialog('modeless', html, title);
  }

  /** @return {string} */
  toString() {
    return 'Ui';
  }

  /**
   * @param {string} kind
   * @param {*} html
   * @param {string=} title
   */
  _recordDialog(kind, html, title) {
    const record = {type: kind, title: title === undefined ? null : String(title), ...describeHtml(html)};
    if (title !== undefined) record.title = String(title);
    this._controller.dialogs.push(record);
    this._controller.history.push(record);
  }
}

module.exports = {UiController, UiMock, MenuMock, PromptResponseMock, BUTTON, BUTTON_SET, defineHidden};
