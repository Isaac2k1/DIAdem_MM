'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 2014-05-22 10:53:18
'-- Ersteller: 
'-- Kommentar: 
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Script.

Dim count_a, count_b

For count_a = 1 to GroupCount
Call GroupDefaultSet(count_a)

'UFD Generell
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_count", "2", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_elFest", "12", DataTypeString)

'UFD1
'--------------------------------------------------------------------------------------------------------------------------------
'Zuordnung O1, O2, C1, C2
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_O1", "UFD1-U-O1 (TCU1)", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_O2", "UFD1-U-O2 (TCU2)", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_C1", "UFD1-U-C1 (TCU3)", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_C2", "UFD1-U-C2 (TCU4)", DataTypeString)

'Zuordnung I-Coil
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_I-Coil", "UFD1-I-Coil", DataTypeString)

'Zuordnung Weg
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_S1", "UFD1-S-DRIVE A", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_S2", "UFD1-S-DRIVE B", DataTypeString)

'Zuordnung PMS
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_PMS1_1", "UFD1-PMS-1.1", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_PMS1_2", "UFD1-PMS-1.2", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_PMS1_3", "UFD1-PMS-1.3", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_PMS1_4", "UFD1-PMS-1.4", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_PMS1_5", "UFD1-PMS-1.5", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_PMS1_6", "UFD1-PMS-1.6", DataTypeString)

Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_PMS2_1", "UFD1-PMS-2.1", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_PMS2_2", "UFD1-PMS-2.2", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_PMS2_3", "UFD1-PMS-2.3", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_PMS2_4", "UFD1-PMS-2.4", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_PMS2_5", "UFD1-PMS-2.5", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_1_PMS2_6", "UFD1-PMS-2.6", DataTypeString)

'UFD2
'--------------------------------------------------------------------------------------------------------------------------------
'Zuordnung O1, O2, C1, C2
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_2_O1", "UFD2-U-O1 (TCU3)", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_2_O2", "UFD2-U-O2 (TCU4)", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_2_C1", "UFD2-U-C1 (TCU1)", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_2_C2", "UFD2-U-C2 (TCU2)", DataTypeString)

'Zuordnung I-Coil
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_2_I-Coil", "UFD2-I-Coil", DataTypeString)

'Zuordnung Weg
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_2_S1", "UFD2-S-DRIVE A", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("UFD_2_S2", "UFD2-S-DRIVE B", DataTypeString)

Next

