using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using BepInEx;
using BepInEx.Logging;
using HarmonyLib;
using NPOI.SS.UserModel;

[BepInPlugin("soeklgb.import_excel_alternative", "ImportExcel Alternative", "1.0.0")]
public class Plugin : BaseUnityPlugin
{
    public void Awake()
    {
        logger = this.Logger;

        var harmony = new Harmony("soeklgb.import_excel_alternative");
        harmony.PatchAll();
    }

    internal static ManualLogSource logger;
}

[HarmonyPatch]
class SourceDataPatch
{
    static IEnumerable<MethodBase> TargetMethods()
    {
        Type[] types =
        [
            typeof(LangGeneral),
            typeof(LangList),
            typeof(LangGame),
            typeof(LangWord),
            typeof(LangNote),
            typeof(LangTalk),
            typeof(SourceArea),
            typeof(SourceBacker),
            typeof(SourceBlock),
            typeof(SourceCalc),
            typeof(SourceCategory),
            typeof(SourceCellEffect),
            typeof(SourceChara),
            typeof(SourceCharaText),
            typeof(SourceCheck),
            typeof(SourceCollectible),
            typeof(SourceElement),
            typeof(SourceFaction),
            typeof(SourceFloor),
            typeof(SourceGlobalTile),
            typeof(SourceHobby),
            typeof(SourceHomeResource),
            typeof(SourceJob),
            typeof(SourceKeyItem),
            typeof(SourceMaterial),
            typeof(SourceObj),
            typeof(SourcePerson),
            typeof(SourceQuest),
            typeof(SourceRace),
            typeof(SourceRecipe),
            typeof(SourceReligion),
            typeof(SourceResearch),
            typeof(SourceSpawnList),
            typeof(SourceStat),
            typeof(SourceTactics),
            typeof(SourceThing),
            typeof(SourceThingV),
            typeof(SourceFood),
            typeof(SourceZone),
            typeof(SourceZoneAffix),
        ];

        return types.Select(t => t.GetMethod("CreateRow")).Cast<MethodBase>();
    }

    static bool Prefix(SourceData __instance, ref object __result)
    {
        Type rowType = __instance.GetType().GetNestedType("Row");
        if (rowType == null)
        {
            return true;
        }

        IRow fieldNameRow = SourceData.row.Sheet.GetRow(0);
        IRow fieldTypeRow = SourceData.row.Sheet.GetRow(1);

        {
            int fieldNameRowLength = fieldNameRow.Where(c => c.StringCellValue != "").Count();
            int fieldTypeRowLength = fieldTypeRow.Where(c => c.StringCellValue != "").Count();
            if (fieldNameRowLength != fieldTypeRowLength)
            {
                return true;
            }
        }

        object row = rowType.GetConstructors().First().Invoke([]);

        for (int index = 0; index < fieldNameRow.Count(); index++)
        {
            string fieldName = fieldNameRow.GetCell(index).StringCellValue;
            string fieldType = fieldTypeRow.GetCell(index).StringCellValue;

            if (fieldName == "")
            {
                continue;
            }

            if (rowType.GetField(fieldName) == null)
            {
                Plugin.logger.LogWarning($"Field \"{fieldName}\" not found in \"{rowType}\"");
                continue;
            }

            switch (fieldType)
            {
                case "int":
                    row.SetField(fieldName, SourceData.GetInt(index));
                    break;
                case "bool":
                    row.SetField(fieldName, SourceData.GetBool(index));
                    break;
                case "double":
                    row.SetField(fieldName, SourceData.GetDouble(index));
                    break;
                case "float":
                    row.SetField(fieldName, SourceData.GetFloat(index));
                    break;
                case "float[]":
                    row.SetField(fieldName, SourceData.GetFloatArray(index));
                    break;
                case "int[]":
                    row.SetField(fieldName, SourceData.GetIntArray(index));
                    break;
                case "string[]":
                    row.SetField(fieldName, SourceData.GetStringArray(index));
                    break;
                case "string":
                    row.SetField(fieldName, SourceData.GetString(index));
                    break;
                case "element":
                case "element_id":
                    row.SetField(fieldName, Core.GetElement(SourceData.GetStr(index, false)));
                    break;
                case "elements":
                    row.SetField(fieldName, Core.ParseElements(SourceData.GetStr(index, false)));
                    break;
                case "":
                    Plugin.logger.LogWarning($"Field \"{fieldName}\" has no type");
                    break;
                default:
                    Plugin.logger.LogWarning(
                        $"Unknown field type \"{fieldType}\" for \"{fieldName}\""
                    );
                    break;
            }
        }

        __result = row;

        return false;
    }
}
