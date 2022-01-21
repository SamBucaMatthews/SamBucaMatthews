import os
import pandas as pd
import xml.etree.ElementTree as et


class Worksheet:
    def __init__(self, file):
        self.tree = et.parse(file)
        self.root = self.tree.getroot()
        self.worksheetCode = self.root.find("basicData/code").text
        self.worksheetVersion = self.root.find(
            "basicData/worksheetVersion").attrib["version"]
        self.other = self.root.find("ingredients/other")
        self.otherItems = self.root.findall("ingredients/other/item")

    def incrementVersion(self):
        newVersion = str(int(self.worksheetVersion) + 1)
        self.worksheetVersion = newVersion
        self.root.find(
            "basicData/worksheetVersion").attrib["version"] = newVersion

    def save(self, filename):
        self.tree.write(filename)


currentWorksheetFolder = r"C:\\Projects\\QA\\Worksheet Updates\\Current"
updatedWorksheetFolder = r"C:\\Projects\\QA\\Worksheet Updates\\Updated"
dodgyStartingMaterialSheet = r"C:\\Projects\\QA\\Worksheet Updates\\StartingMaterialUpdates.xlsx"

os.chdir(currentWorksheetFolder)
namesToUpdate = pd.read_excel(
    dodgyStartingMaterialSheet, sheet_name="Name Updates").set_index("Current Name").transpose().to_dict()

allowedMaterialsToLink = pd.read_excel(
    dodgyStartingMaterialSheet, sheet_name="SMAC Links").transpose().to_dict()


def saveUpdatedVersion(updatedWorksheetFolder, worksheet):
    worksheet.incrementVersion()
    worksheet.save(
        f"{updatedWorksheetFolder}\{worksheet.worksheetCode} v{worksheet.worksheetVersion}.xml")


def updateName(startingMaterialsToUpdate, ingredient, name):
    ingredient.attrib["name"] = startingMaterialsToUpdate[name]["New Name"]


def addAllowedStartingMaterials(startingMaterialsToUpdate, ingredient, name):
    allowedStartingMaterialsElement = et.Element("allowedStartingMaterials")
    ingredient.insert(0, allowedStartingMaterialsElement)

    allowedMaterialIds = [allowedMaterialsToLink[i]["SMAC ID"]
                          for i in allowedMaterialsToLink if allowedMaterialsToLink[i]["Starting Material Name"] == startingMaterialsToUpdate[name]["New Name"]]
    for allowedMaterialId in allowedMaterialIds:
        if allowedMaterialId == "Unsure":
            pass
        else:
            allowedStartingMaterial = et.Element(
                "startingMaterial", id=str(allowedMaterialId))
            allowedStartingMaterialsElement.insert(0, allowedStartingMaterial)


for file in os.listdir():
    try:
        worksheet = Worksheet(file)
        for ingredient in worksheet.otherItems:
            name = ingredient.attrib["name"]
            if namesToUpdate.get(name):
                updateName(namesToUpdate, ingredient, name)
                if ingredient.attrib["name"] == "Remove":
                    worksheet.other.remove(ingredient)
                else:
                    addAllowedStartingMaterials(
                        namesToUpdate, ingredient, name)

        saveUpdatedVersion(updatedWorksheetFolder, worksheet)
    except Exception as e:
        print(worksheet.worksheetCode, e)


def findDodgyOtherStartingMaterials():
    ingredients = []
    os.chdir(currentWorksheetFolder)

    for file in os.listdir():
        try:
            worksheet = Worksheet(file)
            for ingredient in worksheet.otherItems:
                ingredients.append(
                    {
                        "worksheetCode": worksheet.worksheetCode,
                        "ingredient": ingredient.attrib["name"],
                        "hasAllowedStartingMaterials": bool(ingredient.findall("allowedStartingMaterials"))
                    })
        except Exception as e:
            print(worksheet.worksheetCode, e)

    return pd.DataFrame([
        i for i in ingredients if not i["hasAllowedStartingMaterials"]])
