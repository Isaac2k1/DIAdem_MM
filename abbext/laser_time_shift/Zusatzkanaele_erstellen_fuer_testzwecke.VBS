'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 2009-12-09 15:44:08
'-- Ersteller: ---
'-- Kommentar: ---
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Script.

Dim count_a
    
For count_a = 1 to GroupCount
Call GroupDefaultSet(count_a)

Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_count", "2", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_SetIdent", "1", DataTypeString)

Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_sampletime", "40", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_channel_1", "S-UFD1-Drive A", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_channel_2", "S-UFD1-Drive B", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_filter_1", "1", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_filter_2", "1", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_average_1", "1", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_average_2", "1", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_VanHand", "110", DataTypeString)

'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_2_sampletime", "50", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_2_channel_1", "UFD2-Drive A", DataTypeString)
Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_2_channel_2", "S-UFD2-DriveB", DataTypeString)
'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_2_filter_1", "1", DataTypeString)
'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_2_filter_2", "1", DataTypeString)
'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_2_average_1", "1", DataTypeString)
'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_2_average_2", "1", DataTypeString)
'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_2_VanHand", "274", DataTypeString)

'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_3_sampletime", "100", DataTypeString)
'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_3_channel_1", "Laser A", DataTypeString)
'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_3_channel_2", "Laser B", DataTypeString)
'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_3_filter_1", "1", DataTypeString)
'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_3_filter_2", "1", DataTypeString)
'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_3_average_1", "1", DataTypeString)
'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_3_average_2", "1", DataTypeString)
'Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_3_VanHand", "274", DataTypeString)


Next

