//////// Start of button stub functions

/**
 * Roll a d20 using dddice and display the result.
 */
function rollD20() {
  const roll_json = DDDICE_JSON_FROM_ARGS(1, "d20");
  const result = DDDICE_ROLL_FROM_JSON(roll_json);

  let ui = SpreadsheetApp.getUi();
  ui.alert("Notice", `Your dddice roll result was: ${result}`, ui.ButtonSet.OK);
}

//////// End of button stub functions

//////// Start of simple triggers

/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 * @see https://developers.google.com/apps-script/guides/triggers#onopene
 */
function onOpen(e) {
  addThreeDDiceMenu();
}

/**
 * Adds the dddice menu
 */
function addThreeDDiceMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("dddice")
    .addItem("Roll from JSON in active cell", "rollFromActiveCellJSON")
    .addSeparator()
    .addItem("Change the default dice theme", "promptUserForDiceTheme")
    .addItem("Change the default room ID", "promptUserForRoomID")
    .addSeparator()
    .addSubMenu(
      ui.createMenu("Admin actions")
        .addItem("Change the user API key", "promptUserForAuthKey")
        .addSeparator()
        .addItem(
          "Delete user property: user API key",
          "deleteThreeDDiceAuthKey"
        )
        .addItem(
          "Delete user property: default dice theme",
          "deleteDefaultDiceTheme"
        )
        .addItem("Delete user property: default room ID", "deleteDefaultRoomID")
        .addSeparator()
        .addItem("Delete all dddice user properties", "deleteThreeDDiceProps")
    )
    .addToUi();
}

//////// End of simple triggers

//////// Start of installable triggers

/**
 * An event handler which must be installed in the Triggers section.
 * @param {Event} e The onEdit event.
 */
function onEditInstallable(e) {
  // check if event has a range property
  if (e.range) {
    // check if event's range has a size of 1x1
    if (e.range.getNumColumns() === 1 && e.range.getNumRows() === 1) {
      // check if event's range is a checked checkbox
      if (e.range.isChecked()) {
        processThreeDDiceCheckbox(e.range);
      }
    }
  }
}

/**
 * Processes dddice roll from cells by checked checkbox.
 * @param {Range} checkbox The Range of the checked checkbox.
 */
function processThreeDDiceCheckbox(checkbox) {
  const ref_col = checkbox.getColumn();
  const ref_row = checkbox.getRow();
  let sheet = checkbox.getSheet();
  let roll_result = null;
  let roll_direction = null;

  // start by looking to the right
  try {
    let cell_right = sheet.getRange(ref_row, ref_col + 1);
    roll_result = DDDICE_ROLL_FROM_JSON(cell_right.getValue());
    roll_direction = "right";
  } catch {
    // that didn't work. ensure roll_result and roll_direction are null
    roll_result = null;
    roll_direction = null;
  }

  if (roll_result === null) {
    // couldn't roll to the right, so let's try below the checkbox
    try {
      let cell_down = sheet.getRange(ref_row + 1, ref_col);
      roll_result = DDDICE_ROLL_FROM_JSON(cell_down.getValue());
      roll_direction = "down";
    } catch {
      // that didn't work. ensure roll_result and roll_direction are null
      roll_result = null;
      roll_direction = null;
    }
  }

  if (roll_direction !== null) {
    // log the roll result just past the roll's JSON cell
    let log_row = ref_row;
    let log_col = ref_col;
    if (roll_direction === "right") {
      log_col = log_col + 2;
    } else if (roll_direction === "down") {
      log_row = log_row + 2;
    } else {
      Logger.log(`Unknown roll_direction: ${roll_direction}`);
    }

    // check to make sure we've made a change to the log coordinates
    if (log_row !== ref_row || log_col !== ref_col) {
      try {
        let log_cell = sheet.getRange(log_row, log_col);
        log_cell.setValue(roll_result);
      } catch {
        // no log cell to record the result
      }
    }
  }

  // reset the checkbox if we successfully rolled
  if (roll_result !== null) {
    checkbox.uncheck();
  }
}

//////// End of installable triggers

//////// Start of custom functions

/**
 * Formats dice roll arguments into a parseable JSON string for rolling.
 * @param {number} dice_number The number of dice to roll.
 * @param {string} dice_type The type of dice to roll. Must be "d20", "d12", "d10", "d10x", "d8", "d6", "d4", or "mod".
 * @param {number} roll_modifier Optional modifier to the roll.
 * @param {string} dice_theme Optional dice theme ID to use instead of default theme.
 * @param {string} room_id Optional room ID to send roll to instead of default room.
 * @param {string} extra_dice_args_string Optional JSON string to add to die/dice.
 * @param {string} extra_roll_args_string Optional JSON string to add to roll.
 * @returns {string} A JSON string summarizing the roll info.
 * @customfunction
 */
function DDDICE_JSON_FROM_ARGS(
  dice_number,
  dice_type,
  roll_modifier,
  dice_theme,
  room_id,
  extra_dice_args_string,
  extra_roll_args_string
) {
  // define the die/dice for this roll
  let die_to_add = {};
  die_to_add.number = parseInt(dice_number);
  die_to_add.type = dice_type;
  if (dice_theme) die_to_add.theme = dice_theme;
  if (extra_dice_args_string) {
    const extra_dice_args = JSON.parse(extra_dice_args_string);
    die_to_add = addNewKeysFromObjectToObject(die_to_add, extra_dice_args);
  }

  // define the JSON for this roll
  let roll_json = {
    dice_to_add: [die_to_add],
  };
  if (roll_modifier) roll_json.roll_modifier = parseInt(roll_modifier);
  if (room_id) roll_json.room = room_id;
  if (extra_roll_args_string) {
    const extra_roll_args = JSON.parse(extra_roll_args_string);
    roll_json = addNewKeysFromObjectToObject(roll_json, extra_roll_args);
  }

  // add a key for readability in the cell
  let output = {
    dddice_roll: roll_json,
  };
  return JSON.stringify(output);
}

/**
 * Combines multiple dddice roll JSONs into one roll.
 * @param {string[]} json_strings The JSON strings to combine.
 * @returns {string} The combined dddice JSON string.
 * @customfunction
 */
function COMBINE_DDDICE_JSONS(...json_strings) {
  let roll_json = {
    dice_to_add: [],
  };
  for (let i = 0; i < json_strings.length; i++) {
    let parsed_json;
    try {
      parsed_json = JSON.parse(json_strings[i]);
    } catch {
      continue;
    }
    let roll_i = parsed_json.hasOwnProperty("dddice_roll")
      ? parsed_json["dddice_roll"]
      : parsed_json;

    if (roll_i.hasOwnProperty("dice")) {
      if (!roll_json.hasOwnProperty("dice")) {
        roll_json.dice = [];
      }
      roll_json.dice.push(...roll_i.dice);
    }

    if (roll_i.hasOwnProperty("dice_to_add")) {
      roll_json.dice_to_add.push(...roll_i.dice_to_add);
    }

    if (roll_i.hasOwnProperty("roll_modifier")) {
      if (!roll_json.hasOwnProperty("roll_modifier")) {
        roll_json.roll_modifier = 0;
      }
      roll_json.roll_modifier = roll_json.roll_modifier + roll_i.roll_modifier;
    }

    // add the rest of the keys, ignoring new value if key already exists
    roll_json = addNewKeysFromObjectToObject(roll_json, roll_i);
  }

  // add a key for readability in the cell
  let output = {
    dddice_roll: roll_json,
  };
  return JSON.stringify(output);
}

//////// End of custom functions

/**
 * Adds key-value pairs to an object if not yet defined.
 * @param {Object} target_object Object to add new keys to.
 * @param {Object} object_to_add Object adding keys from.
 * @returns {Object} New combined object.
 */
function addNewKeysFromObjectToObject(target_object, object_to_add) {
  const current_keys = Object.keys(target_object);
  const keys_to_add = Object.keys(object_to_add);
  for (let i = 0; i < keys_to_add.length; i++) {
    let key_i = keys_to_add[i];
    if (!current_keys.includes(key_i)) {
      target_object[key_i] = object_to_add[key_i];
    }
  }
  return target_object;
}

/**
 * Rolls selected JSON and displays the value
 */
function rollFromActiveCellJSON() {
  const result = DDDICE_ROLL_FROM_JSON(
    SpreadsheetApp.getActiveRange().getValue()
  );

  let ui = SpreadsheetApp.getUi();
  ui.alert("Notice", `Your dddice roll result was: ${result}`, ui.ButtonSet.OK);
}

/**
 * Executes a dddice roll defined by the input JSON string.
 * @param {string} json_string JSON string defining dddice roll.
 * @returns {number} The total value of the roll.
 */
function DDDICE_ROLL_FROM_JSON(json_string) {
  const url = "https://dddice.com/api/1.0/roll";

  const auth_key = getThreeDDiceAuthKey();
  const headers = {
    Authorization: `Bearer ${auth_key}`,
    "Content-Type": "application/json",
    Accept: "application/json",
  };
  const body = convertJSONStringToBody(json_string);

  let response = UrlFetchApp.fetch(url, {
    method: "POST",
    headers: headers,
    payload: JSON.stringify(body),
  });
  const response_json = JSON.parse(response.getContentText());
  return response_json.data.total_value;
}

/**
 * Converts the summary JSON string into the body format expected by the dddice API
 * @param {string} json_string JSON string defining a dddice roll.
 * @returns {Object} The body object, formatted in the way the dddice API expects.
 */
function convertJSONStringToBody(json_string) {
  if (json_string === undefined)
    json_string = '{"dice_to_add":[{"number":1,"type":"d20"}]}';
  const parsed_json = JSON.parse(json_string);
  let roll_json = parsed_json.hasOwnProperty("dddice_roll")
    ? parsed_json["dddice_roll"]
    : parsed_json;

  // create the dice array, unless it already exists
  if (!roll_json.hasOwnProperty("dice")) roll_json.dice = [];

  // add all the dice described in the dice_to_add array
  if (roll_json.hasOwnProperty("dice_to_add")) {
    roll_json.dice.push(...convertDiceToAddArray(roll_json.dice_to_add));
    delete roll_json.dice_to_add; // clean up afterwards
  }

  // add the roll modifier if provided
  if (roll_json.hasOwnProperty("roll_modifier")) {
    roll_json.dice.push({
      type: "mod",
      value: roll_json.roll_modifier,
    });
    delete roll_json.roll_modifier; // clean up afterwards
  }

  // add the default dice theme to all dice which don't already have a theme
  const default_dice_theme = getDefaultDiceTheme();
  for (let i = 0; i < roll_json.dice.length; i++) {
    if (!roll_json.dice[i].hasOwnProperty("theme")) {
      roll_json.dice[i].theme = default_dice_theme;
    }
  }

  // add the default room ID if none already specified
  if (!roll_json.hasOwnProperty("room")) {
    const default_room_id = getDefaultRoomID();
    if (default_room_id) roll_json.room = default_room_id;
  }

  return roll_json;
}

/**
 * Converts the summary array of dice to the API-style dice array.
 * @param {Array} dice_to_add Summary array of dice to be converted.
 * @returns {Array} The dice array in the API format.
 */
function convertDiceToAddArray(dice_to_add) {
  let output = [];
  for (let i = 0; i < dice_to_add.length; i++) {
    let die_to_add = {
      type: dice_to_add[i].type,
    };
    if (dice_to_add[i].hasOwnProperty("theme")) {
      die_to_add.theme = dice_to_add[i].theme;
    }
    let number_to_add = dice_to_add[i].hasOwnProperty("number")
      ? dice_to_add[i].number
      : 1;
    for (let j = 0; j < number_to_add; j++) {
      output.push(die_to_add);
    }
  }
  return output;
}

/**
 * Gets the dddice user API key, prompting the user if not yet defined.
 * @returns {string} The dddice user API key.
 */
function getThreeDDiceAuthKey() {
  let user_props = PropertiesService.getUserProperties();
  let dddice_auth_key = user_props.getProperty("dddice_auth_key");
  if (dddice_auth_key === null) {
    return promptUserForAuthKey();
  }
  return dddice_auth_key;
}

/**
 * Gets the default dice theme, prompting the user if not yet defined.
 * @returns {string} The default dddice dice theme.
 */
function getDefaultDiceTheme() {
  let user_props = PropertiesService.getUserProperties();
  let dddice_default_dice_theme = user_props.getProperty(
    "dddice_default_dice_theme"
  );
  if (dddice_default_dice_theme === null) {
    return promptUserForDiceTheme();
  }
  return dddice_default_dice_theme;
}

/**
 * Gets the default room ID, prompting the user if not yet defined.
 * @returns {string} The default dddice room ID.
 */
function getDefaultRoomID() {
  let user_props = PropertiesService.getUserProperties();
  let dddice_default_room_id = user_props.getProperty("dddice_default_room_id");
  if (dddice_default_room_id === null) {
    return promptUserForRoomID();
  }
  return dddice_default_room_id;
}

/**
 * Prompts the user for their dddice API key. If provided, stores as a user property.
 * @returns {string|null} The user API key or null if not provided.
 */
function promptUserForAuthKey() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.prompt(
    "User property input",
    "Please provide your user API key for dddice.",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() != ui.Button.OK) return null;
  const auth_key = response.getResponseText();
  PropertiesService.getUserProperties().setProperty(
    "dddice_auth_key",
    auth_key
  );
  return auth_key;
}

/**
 * Prompts the user for a default dice theme. Stores result as a user property.
 * @returns {string} The default dice theme. Uses "dddice-red" if not provided.
 */
function promptUserForDiceTheme() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.prompt(
    "User property input",
    "Please provide a default dice theme for dddice rolls.",
    ui.ButtonSet.OK_CANCEL
  );
  let dice_theme = response.getResponseText();
  if (!dice_theme) {
    const fallback_dice_theme = "dddice-red";
    dice_theme = fallback_dice_theme;
  }
  PropertiesService.getUserProperties().setProperty(
    "dddice_default_dice_theme",
    dice_theme
  );
  return dice_theme;
}

/**
 * Prompts the user for a default room ID. Stores result as a user property.
 * @returns {string} The default room ID. Uses "" if not provided.
 */
function promptUserForRoomID() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.prompt(
    "User property input",
    "Please provide a default room ID for dddice rolls.",
    ui.ButtonSet.OK_CANCEL
  );
  let room_id = response.getResponseText();
  if (!room_id) {
    const fallback_room_id = "";
    room_id = fallback_room_id;
  }
  PropertiesService.getUserProperties().setProperty(
    "dddice_default_room_id",
    room_id
  );
  return room_id;
}

/**
 * Deletes all the dddice user properties after confirmation.
 */
function deleteThreeDDiceProps() {
  let ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Confirm deletion",
    "Are you sure you want to delete ALL the dddice user properties?",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) {
    return;
  }

  const confirmed = true;
  deleteThreeDDiceAuthKey(confirmed);
  deleteDefaultDiceTheme(confirmed);
  deleteDefaultRoomID(confirmed);
}

/**
 * Deletes the dddice API key from the user properties after confirmation.
 * @param {bool} confirmed If true, skips user confirmation.
 */
function deleteThreeDDiceAuthKey(confirmed) {
  if (!confirmed) {
    let ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "Confirm deletion",
      "Are you sure you want to delete the user property: user API key?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return;
    }
  }

  PropertiesService.getUserProperties().deleteProperty("dddice_auth_key");
}

/**
 * Deletes the default dice theme from the user properties after confirmation.
 * @param {bool} confirmed If true, skips user confirmation.
 */
function deleteDefaultDiceTheme(confirmed) {
  if (!confirmed) {
    let ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "Confirm deletion",
      "Are you sure you want to delete the user property: default dice theme?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return;
    }
  }

  PropertiesService.getUserProperties().deleteProperty(
    "dddice_default_dice_theme"
  );
}

/**
 * Deletes the default room ID from the user properties after confirmation.
 * @param {bool} confirmed If true, skips user confirmation.
 */
function deleteDefaultRoomID(confirmed) {
  if (!confirmed) {
    let ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "Confirm deletion",
      "Are you sure you want to delete the user property: default room ID?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return;
    }
  }

  PropertiesService.getUserProperties().deleteProperty(
    "dddice_default_room_id"
  );
}
