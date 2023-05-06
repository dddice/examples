# dddice Google Sheets Example

An example of how to use the dddice API in Google Sheets using Google Apps Script.

## Usage

To use this example, you'll have to create a Google Apps Script attached to a spreadsheet and authorize it. To do so:

1. Open the Google Sheets spreadsheet you want to roll from.
2. Open the **Apps Script** editor in the **Extensions** menu.
3. Give your new Apps Script project a name, such as "dddice roller"
4. Copy-paste the contents of the `Code.js` file into the editor (replacing the empty `myFunction` which is there) and save.
5. At the top of the editor, click "Run" to execute the `rollD20` function.
6. An "Authorization required" popup will open. Select "Review permissions" from it.
7. Choose your Google account if prompted to do so.
8. At the "Google hasn't verified this app" screen, select "Advanced" in the lower-left corner, followed by "Go to dddice roller (unsafe)".
9. Unfortunately, the permissions are rather scary looking even though the script will only edit the spreadsheet you're in. Click "Allow" to continue.
10. Return to the spreadsheet, and you'll see a popup titled "User property input" which is asking for your dddice API key. You need one to continue, so go to [dddice.com/account/developer](https://dddice.com/account/developer)
11. Click "Create API Key" which will generate a key and copy it to your clipboard.
12. Return to the spreadsheet and paste the API key into the popup's input. Click "Ok" to continue.
13. You will now see a popup asking for a default dice theme to use for dddice rolls; a dice theme's ID is the last part of the URL of its [dice page](https://dddice.com/dice). If you're not sure what theme to use, click "Ok" and it will default to "dddice-red". 
14. Next, you'll be asked for a default room ID for your dddice rolls. If you have a room you want to roll to, copy-paste its ID (the last part of the room's URL). Otherwise, you can leave it blank and click "Ok" to continue with no room.
15. Finally, you'll get a popup titled "Notice" with the result of your d20 roll. If you put in a room ID in the last step, you should see the roll in there too!

### The dddice Menu: Editing the User Properties

If you want to edit the API key, the default dice theme, or the default room ID, it's easy to do so using the **dddice** menu which the script adds to your spreadsheet. If you don't see it, refresh your spreadsheet (this will close the Apps Script editor) or run the `onOpen` function.

The **dddice** menu has several options:

- **Roll from JSON in active cell**
  - Highlight a cell containing the JSON string result of the `DDDICE_JSON_FROM_ARGS` custom function and select this option to roll it.
- **Change the default dice theme**
  - Selecting this prompts you to enter a new default dice theme.
- **Change the default room ID**
  - Selecting this prompts you to enter a new default room ID.
- **Admin actions**
  - These are actions you're probably going to need less often. Using these, you can change the user API key or delete any/all of the user properties stored by the script.

### Custom Functions

The script adds two new custom functions for you to use:

- **`DDDICE_JSON_FROM_ARGS`**
  - This function lets you specify a dddice roll: you can roll 1 or more of the same type of die, add a modifier, specify a different theme or room ID, and you can add additional API parameters as a JSON string.
  - The output of this function is just a JSON string; it doesn't actually trigger the roll. To do so, highlight the cell and select **Roll from JSON in active cell** in the **dddice** menu.
- **`COMBINE_DDDICE_JSONS`**
  - This function lets you combine multiple dddice JSON strings into one so you can do more complex rolls involving multiple types of dice.
  - Just like with the other function, the output of this is just a JSON string which you can roll from using the **dddice** menu.

As an example, let's say we have the following:

- In cell **B2**: `=DDDICE_JSON_FROM_ARGS(1, "d20", 7)`
  - This would roll `1d20 + 7` with the default dice theme in the default room.
- In cell **B3**: `=DDDICE_JSON_FROM_ARGS(1, "d8", 4)`
  - This would roll `1d8 + 4` with the default dice theme in the default room.
- In cell **B4**: `=DDDICE_JSON_FROM_ARGS(3, "d6", 0, "dddice-black", , "{""label"": ""Sneak Attack""}")`
  - This would roll `3d6` with the "dddice-black" theme in the default room, with the added label of "Sneak Attack".
- In cell **B5**: `=COMBINE_DDDICE_JSONS(B3, B4)`
  - This would roll the two above rolls at the same time.

### Activating with Checkboxes

Opening the **dddice** menu for every roll is a lot of clicks, so I've added an option to activate dice rolls with checkboxes. To do so, you need to install a trigger for a function.

1. Open the Apps Script editor.
2. Open the "Triggers" section from the left-side menu.
3. In the bottom-right corner, click the "Add Trigger" button.
4. Under "Choose which function to run", select the `onEditInstallable` function from the dropdown.
5. Under "Select event type", select "On edit" from the dropdown.
6. Click "Save".

Now, whenever you edit the spreadsheet, the `onEditInstallable` function will be called. That function checks if you've just checked a checkbox, and if you have, it will call the `processThreeDDiceCheckbox` function. That function looks to the right of the checkbox for a dddice roll JSON, and if it can't find one there, it looks under the checkbox. If it successfully rolls from one of those neighboring cells, it will try to put the result of that roll into the cell that's just past the JSON cell. Finally, if it was able to do a roll, the script will uncheck the checkbox.

To try it out, first populate cells **B2:B5** as described in the example above. Then highlight cells **A2:A5** and click **Insert** > **Checkbox** from the top menu. Now if you check one of those checkboxes, the roll right next to it will be executed, and the result will be put into column **C**. The checkbox then gets unchecked so you can click it again.

### Activating with Buttons

There's another option for activating dddice rolls: making a "button" which is a drawing with an attached script. This option will be most useful to you if you're comfortable adding new functions in the Apps Script editor. JavaScript experience is helpful.

To add a button, do the following:

1. In the spreadsheet, click **Insert** > **Drawing** from the top menu.
2. In the drawing editor popup, make a drawing. For example, click **Shape** > **Shapes** > **Rectangle** from the menu and draw it in.
3. Click "Save and Close" in the upper-right corner.
4. Your drawing should now be in the spreadsheet, above the cells. Move and resize it until you're satisfied with its position.
5. With the drawing selected, three dots should appear in its upper-right corner. Click those dots and select "Assign script" from the menu that it opens.
6. In the "Assign script" popup, type in the name of the function you want to use. For example, you could type in `rollD20` or `rollFromActiveCellJSON`. Click "OK".
7. You've now set up your drawing as a button! Click it to activate the function.

If you want to make any changes to the button (e.g. move/resize it, change the function, delete it), you'll have to `right-click` or `Ctrl/Cmd + left-click` it to select it and not activate the function.

There's a section at the top of the script for "button stub functions" with an example function: `rollD20`. You can add your own button functions, using that function as an example. If you want to make reference to particular cell values within a function, that will require more work on your part; I suggest just using `DDDICE_JSON_FROM_ARGS` and `COMBINE_DDDICE_JSONS` to roll while referencing cell values.

## Credits

This example was written in response to [an integration request](https://discord.com/channels/905202772482359327/1074475474916474950/1074475474916474950) on the dddice Discord server. This work is based on [an example function](https://discord.com/channels/905202772482359327/1074475474916474950/1100183389228761228) written by another user in that thread.
