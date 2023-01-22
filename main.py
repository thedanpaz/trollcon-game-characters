from openpyxl import Workbook

CHARACTER_TYPES = ["scoundrel", "brute"]

def createCharacterSheet(characterName, characterClass, characterLevel):

    # Ensure the characterClass the user is trying to make exists
    if not (characterClass.lower() in CHARACTER_TYPES):
        str = " "
        raise Exception("Character Class: " + characterClass + " is not valid. Must be of type: " + str.join(CHARACTER_TYPES))

    # Instantiate character specifics
    character = {
        "name": characterName,
        "class": characterClass,
        "level": characterLevel,
        "health": 0,
        "strength": 0,
        "dexterity": 0
    }

    match character["class"].lower():
        case "brute":
            character["health"] = 12 + (3 * character["level"])
            character["strength"] = 2 + character["level"]

        case "scoundrel":
            character["health"] = 8 + (2 * character["level"])
            character["dexterity"] = 1 + (2 * character["level"])

    return character


# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    wb = Workbook()
    ws = wb.active
    ws.title = "Gloomhaven"

    # set spreadsheet headers
    ws['A1'] = "Name"
    ws['B1'] = "Class"
    ws['C1'] = "Level"
    ws['D1'] = "Health"
    ws['E1'] = "Strength"
    ws['F1'] = "Dexterity"

    characterList = []
    try:
        characterList.append(createCharacterSheet('jarg', 'Scoundrel', 69))
        characterList.append(createCharacterSheet('Dunbar', 'Brute', 17))
        characterList.append(createCharacterSheet('Dunbar', 'DickButt', 12))
    except Exception as error:
        print(error)

    rowNumber = 1

    # Start appending Character info to spreadsheet
    for character in characterList:
        print(character)
        rowNumber = rowNumber + 1
        # Traits
        ws['A' + str(rowNumber)] = character["name"]
        ws['B' + str(rowNumber)] = character["class"]
        ws['C' + str(rowNumber)] = character["level"]
        ws['D' + str(rowNumber)] = character["health"]
        ws['E' + str(rowNumber)] = character["strength"]
        ws['F' + str(rowNumber)] = character["dexterity"]

    wb.save('/var/www/Character-Sheet.xlsx')
