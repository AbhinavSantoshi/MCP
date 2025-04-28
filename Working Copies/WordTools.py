from typing import Any
from mcp.server.fastmcp import FastMCP
import win32com.client as win32
import uuid

WordObjectIdKey = "__WordObjId"

# Initialize FastMCP server
mcp = FastMCP("word")

# Global Word instance to be ensured by EnsureWord
wordInstance = None

# An object cache meant to keep track of objects being returned
object_cache = {}
def add_object(obj):
    guid = str(uuid.uuid4())
    object_cache[guid] = obj
    return guid

def object_exists(guid):
    return guid in object_cache

def get_object(guid):
    if guid not in object_cache:
        raise ValueError(f"Value {guid} is not the GUID of a known object. To pass objects, when calling tools, get the object from Word and send it is {WordObjectIdKey}.")
    return object_cache.get(guid)

def remove_object(guid):
    if guid in object_cache:
        del object_cache[guid]

def EnsureWord():
    global wordInstance
    if (wordInstance is None):
        wordInstance = win32.gencache.EnsureDispatch('Word.Application')

    return wordInstance


def tryParseString(value):
    # Try to parse as boolean
    if value.lower() in ['true', 'false']:
        return value.lower() == 'true'

    # Try to parse as integer
    try:
        return int(value)
    except ValueError:
        pass

    # Keep as string
    return value


# Tool: 1
@mcp.tool()
async def word_get_Name():
	this_Global = EnsureWord()
	retVal = this_Global.Name
	return retVal


# Tool: 2
@mcp.tool()
async def word_get_Documents():
	this_Global = EnsureWord()
	retVal = this_Global.Documents
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Documents", "Count": local_Count, }


# Tool: 3
@mcp.tool()
async def word_get_Windows():
	this_Global = EnsureWord()
	retVal = this_Global.Windows
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	try:
		local_SyncScrollingSideBySide = retVal.SyncScrollingSideBySide
	except:
		local_SyncScrollingSideBySide = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Windows", "Count": local_Count, "SyncScrollingSideBySide": local_SyncScrollingSideBySide, }


# Tool: 4
@mcp.tool()
async def word_get_ActiveDocument():
	this_Global = EnsureWord()
	retVal = this_Global.ActiveDocument
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Document"}


# Tool: 5
@mcp.tool()
async def word_get_ActiveWindow():
	this_Global = EnsureWord()
	retVal = this_Global.ActiveWindow
	try:
		local_Left = retVal.Left
	except:
		local_Left = None
	try:
		local_Top = retVal.Top
	except:
		local_Top = None
	try:
		local_Width = retVal.Width
	except:
		local_Width = None
	try:
		local_Height = retVal.Height
	except:
		local_Height = None
	try:
		local_Split = retVal.Split
	except:
		local_Split = None
	try:
		local_SplitVertical = retVal.SplitVertical
	except:
		local_SplitVertical = None
	try:
		local_Caption = retVal.Caption
	except:
		local_Caption = None
	try:
		local_WindowState = retVal.WindowState
	except:
		local_WindowState = None
	try:
		local_DisplayRulers = retVal.DisplayRulers
	except:
		local_DisplayRulers = None
	try:
		local_DisplayVerticalRuler = retVal.DisplayVerticalRuler
	except:
		local_DisplayVerticalRuler = None
	try:
		local_Type = retVal.Type
	except:
		local_Type = None
	try:
		local_WindowNumber = retVal.WindowNumber
	except:
		local_WindowNumber = None
	try:
		local_DisplayVerticalScrollBar = retVal.DisplayVerticalScrollBar
	except:
		local_DisplayVerticalScrollBar = None
	try:
		local_DisplayHorizontalScrollBar = retVal.DisplayHorizontalScrollBar
	except:
		local_DisplayHorizontalScrollBar = None
	try:
		local_StyleAreaWidth = retVal.StyleAreaWidth
	except:
		local_StyleAreaWidth = None
	try:
		local_DisplayScreenTips = retVal.DisplayScreenTips
	except:
		local_DisplayScreenTips = None
	try:
		local_HorizontalPercentScrolled = retVal.HorizontalPercentScrolled
	except:
		local_HorizontalPercentScrolled = None
	try:
		local_VerticalPercentScrolled = retVal.VerticalPercentScrolled
	except:
		local_VerticalPercentScrolled = None
	try:
		local_DocumentMap = retVal.DocumentMap
	except:
		local_DocumentMap = None
	try:
		local_Active = retVal.Active
	except:
		local_Active = None
	try:
		local_DocumentMapPercentWidth = retVal.DocumentMapPercentWidth
	except:
		local_DocumentMapPercentWidth = None
	try:
		local_Index = retVal.Index
	except:
		local_Index = None
	try:
		local_IMEMode = retVal.IMEMode
	except:
		local_IMEMode = None
	try:
		local_UsableWidth = retVal.UsableWidth
	except:
		local_UsableWidth = None
	try:
		local_UsableHeight = retVal.UsableHeight
	except:
		local_UsableHeight = None
	try:
		local_EnvelopeVisible = retVal.EnvelopeVisible
	except:
		local_EnvelopeVisible = None
	try:
		local_DisplayRightRuler = retVal.DisplayRightRuler
	except:
		local_DisplayRightRuler = None
	try:
		local_DisplayLeftScrollBar = retVal.DisplayLeftScrollBar
	except:
		local_DisplayLeftScrollBar = None
	try:
		local_Visible = retVal.Visible
	except:
		local_Visible = None
	try:
		local_Thumbnails = retVal.Thumbnails
	except:
		local_Thumbnails = None
	try:
		local_ShowSourceDocuments = retVal.ShowSourceDocuments
	except:
		local_ShowSourceDocuments = None
	try:
		local_Hwnd = retVal.Hwnd
	except:
		local_Hwnd = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Window", "Left": local_Left, "Top": local_Top, "Width": local_Width, "Height": local_Height, "Split": local_Split, "SplitVertical": local_SplitVertical, "Caption": local_Caption, "WindowState": local_WindowState, "DisplayRulers": local_DisplayRulers, "DisplayVerticalRuler": local_DisplayVerticalRuler, "Type": local_Type, "WindowNumber": local_WindowNumber, "DisplayVerticalScrollBar": local_DisplayVerticalScrollBar, "DisplayHorizontalScrollBar": local_DisplayHorizontalScrollBar, "StyleAreaWidth": local_StyleAreaWidth, "DisplayScreenTips": local_DisplayScreenTips, "HorizontalPercentScrolled": local_HorizontalPercentScrolled, "VerticalPercentScrolled": local_VerticalPercentScrolled, "DocumentMap": local_DocumentMap, "Active": local_Active, "DocumentMapPercentWidth": local_DocumentMapPercentWidth, "Index": local_Index, "IMEMode": local_IMEMode, "UsableWidth": local_UsableWidth, "UsableHeight": local_UsableHeight, "EnvelopeVisible": local_EnvelopeVisible, "DisplayRightRuler": local_DisplayRightRuler, "DisplayLeftScrollBar": local_DisplayLeftScrollBar, "Visible": local_Visible, "Thumbnails": local_Thumbnails, "ShowSourceDocuments": local_ShowSourceDocuments, "Hwnd": local_Hwnd, }


# Tool: 6
@mcp.tool()
async def word_get_Selection():
	this_Global = EnsureWord()
	retVal = this_Global.Selection
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_Type = retVal.Type
	except:
		local_Type = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Flags = retVal.Flags
	except:
		local_Flags = None
	try:
		local_Active = retVal.Active
	except:
		local_Active = None
	try:
		local_StartIsActive = retVal.StartIsActive
	except:
		local_StartIsActive = None
	try:
		local_IPAtEndOfLine = retVal.IPAtEndOfLine
	except:
		local_IPAtEndOfLine = None
	try:
		local_ExtendMode = retVal.ExtendMode
	except:
		local_ExtendMode = None
	try:
		local_ColumnSelectMode = retVal.ColumnSelectMode
	except:
		local_ColumnSelectMode = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HasChildShapeRange = retVal.HasChildShapeRange
	except:
		local_HasChildShapeRange = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Selection", "Text": local_Text, "Start": local_Start, "End": local_End, "Type": local_Type, "StoryType": local_StoryType, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Information": local_Information, "Flags": local_Flags, "Active": local_Active, "StartIsActive": local_StartIsActive, "IPAtEndOfLine": local_IPAtEndOfLine, "ExtendMode": local_ExtendMode, "ColumnSelectMode": local_ColumnSelectMode, "Orientation": local_Orientation, "NoProofing": local_NoProofing, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HasChildShapeRange": local_HasChildShapeRange, }


# Tool: 7
@mcp.tool()
async def word_get_WordBasic():
	this_Global = EnsureWord()
	retVal = this_Global.WordBasic
	return retVal


# Tool: 8
@mcp.tool()
async def word_get_PrintPreview():
	this_Global = EnsureWord()
	retVal = this_Global.PrintPreview
	return retVal


# Tool: 9
@mcp.tool()
async def word_get_RecentFiles():
	this_Global = EnsureWord()
	retVal = this_Global.RecentFiles
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	try:
		local_Maximum = retVal.Maximum
	except:
		local_Maximum = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "RecentFiles", "Count": local_Count, "Maximum": local_Maximum, }


# Tool: 10
@mcp.tool()
async def word_get_NormalTemplate():
	this_Global = EnsureWord()
	retVal = this_Global.NormalTemplate
	try:
		local_Name = retVal.Name
	except:
		local_Name = None
	try:
		local_Path = retVal.Path
	except:
		local_Path = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_Saved = retVal.Saved
	except:
		local_Saved = None
	try:
		local_Type = retVal.Type
	except:
		local_Type = None
	try:
		local_FullName = retVal.FullName
	except:
		local_FullName = None
	try:
		local_BuiltInDocumentProperties = retVal.BuiltInDocumentProperties
	except:
		local_BuiltInDocumentProperties = None
	try:
		local_CustomDocumentProperties = retVal.CustomDocumentProperties
	except:
		local_CustomDocumentProperties = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_KerningByAlgorithm = retVal.KerningByAlgorithm
	except:
		local_KerningByAlgorithm = None
	try:
		local_JustificationMode = retVal.JustificationMode
	except:
		local_JustificationMode = None
	try:
		local_FarEastLineBreakLevel = retVal.FarEastLineBreakLevel
	except:
		local_FarEastLineBreakLevel = None
	try:
		local_NoLineBreakBefore = retVal.NoLineBreakBefore
	except:
		local_NoLineBreakBefore = None
	try:
		local_NoLineBreakAfter = retVal.NoLineBreakAfter
	except:
		local_NoLineBreakAfter = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_FarEastLineBreakLanguage = retVal.FarEastLineBreakLanguage
	except:
		local_FarEastLineBreakLanguage = None
	try:
		local_AutoSaveOn = retVal.AutoSaveOn
	except:
		local_AutoSaveOn = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Template", "Name": local_Name, "Path": local_Path, "LanguageID": local_LanguageID, "Saved": local_Saved, "Type": local_Type, "FullName": local_FullName, "BuiltInDocumentProperties": local_BuiltInDocumentProperties, "CustomDocumentProperties": local_CustomDocumentProperties, "LanguageIDFarEast": local_LanguageIDFarEast, "KerningByAlgorithm": local_KerningByAlgorithm, "JustificationMode": local_JustificationMode, "FarEastLineBreakLevel": local_FarEastLineBreakLevel, "NoLineBreakBefore": local_NoLineBreakBefore, "NoLineBreakAfter": local_NoLineBreakAfter, "NoProofing": local_NoProofing, "FarEastLineBreakLanguage": local_FarEastLineBreakLanguage, "AutoSaveOn": local_AutoSaveOn, }


# Tool: 11
@mcp.tool()
async def word_get_System():
	this_Global = EnsureWord()
	retVal = this_Global.System
	try:
		local_OperatingSystem = retVal.OperatingSystem
	except:
		local_OperatingSystem = None
	try:
		local_ProcessorType = retVal.ProcessorType
	except:
		local_ProcessorType = None
	try:
		local_Version = retVal.Version
	except:
		local_Version = None
	try:
		local_FreeDiskSpace = retVal.FreeDiskSpace
	except:
		local_FreeDiskSpace = None
	try:
		local_Country = retVal.Country
	except:
		local_Country = None
	try:
		local_LanguageDesignation = retVal.LanguageDesignation
	except:
		local_LanguageDesignation = None
	try:
		local_HorizontalResolution = retVal.HorizontalResolution
	except:
		local_HorizontalResolution = None
	try:
		local_VerticalResolution = retVal.VerticalResolution
	except:
		local_VerticalResolution = None
	try:
		local_ProfileString = retVal.ProfileString
	except:
		local_ProfileString = None
	try:
		local_PrivateProfileString = retVal.PrivateProfileString
	except:
		local_PrivateProfileString = None
	try:
		local_MathCoprocessorInstalled = retVal.MathCoprocessorInstalled
	except:
		local_MathCoprocessorInstalled = None
	try:
		local_ComputerType = retVal.ComputerType
	except:
		local_ComputerType = None
	try:
		local_MacintoshName = retVal.MacintoshName
	except:
		local_MacintoshName = None
	try:
		local_QuickDrawInstalled = retVal.QuickDrawInstalled
	except:
		local_QuickDrawInstalled = None
	try:
		local_Cursor = retVal.Cursor
	except:
		local_Cursor = None
	try:
		local_CountryRegion = retVal.CountryRegion
	except:
		local_CountryRegion = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "System", "OperatingSystem": local_OperatingSystem, "ProcessorType": local_ProcessorType, "Version": local_Version, "FreeDiskSpace": local_FreeDiskSpace, "Country": local_Country, "LanguageDesignation": local_LanguageDesignation, "HorizontalResolution": local_HorizontalResolution, "VerticalResolution": local_VerticalResolution, "ProfileString": local_ProfileString, "PrivateProfileString": local_PrivateProfileString, "MathCoprocessorInstalled": local_MathCoprocessorInstalled, "ComputerType": local_ComputerType, "MacintoshName": local_MacintoshName, "QuickDrawInstalled": local_QuickDrawInstalled, "Cursor": local_Cursor, "CountryRegion": local_CountryRegion, }


# Tool: 12
@mcp.tool()
async def word_get_AutoCorrect():
	this_Global = EnsureWord()
	retVal = this_Global.AutoCorrect
	try:
		local_CorrectDays = retVal.CorrectDays
	except:
		local_CorrectDays = None
	try:
		local_CorrectInitialCaps = retVal.CorrectInitialCaps
	except:
		local_CorrectInitialCaps = None
	try:
		local_CorrectSentenceCaps = retVal.CorrectSentenceCaps
	except:
		local_CorrectSentenceCaps = None
	try:
		local_ReplaceText = retVal.ReplaceText
	except:
		local_ReplaceText = None
	try:
		local_FirstLetterAutoAdd = retVal.FirstLetterAutoAdd
	except:
		local_FirstLetterAutoAdd = None
	try:
		local_TwoInitialCapsAutoAdd = retVal.TwoInitialCapsAutoAdd
	except:
		local_TwoInitialCapsAutoAdd = None
	try:
		local_CorrectCapsLock = retVal.CorrectCapsLock
	except:
		local_CorrectCapsLock = None
	try:
		local_CorrectHangulAndAlphabet = retVal.CorrectHangulAndAlphabet
	except:
		local_CorrectHangulAndAlphabet = None
	try:
		local_HangulAndAlphabetAutoAdd = retVal.HangulAndAlphabetAutoAdd
	except:
		local_HangulAndAlphabetAutoAdd = None
	try:
		local_ReplaceTextFromSpellingChecker = retVal.ReplaceTextFromSpellingChecker
	except:
		local_ReplaceTextFromSpellingChecker = None
	try:
		local_OtherCorrectionsAutoAdd = retVal.OtherCorrectionsAutoAdd
	except:
		local_OtherCorrectionsAutoAdd = None
	try:
		local_CorrectKeyboardSetting = retVal.CorrectKeyboardSetting
	except:
		local_CorrectKeyboardSetting = None
	try:
		local_CorrectTableCells = retVal.CorrectTableCells
	except:
		local_CorrectTableCells = None
	try:
		local_DisplayAutoCorrectOptions = retVal.DisplayAutoCorrectOptions
	except:
		local_DisplayAutoCorrectOptions = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "AutoCorrect", "CorrectDays": local_CorrectDays, "CorrectInitialCaps": local_CorrectInitialCaps, "CorrectSentenceCaps": local_CorrectSentenceCaps, "ReplaceText": local_ReplaceText, "FirstLetterAutoAdd": local_FirstLetterAutoAdd, "TwoInitialCapsAutoAdd": local_TwoInitialCapsAutoAdd, "CorrectCapsLock": local_CorrectCapsLock, "CorrectHangulAndAlphabet": local_CorrectHangulAndAlphabet, "HangulAndAlphabetAutoAdd": local_HangulAndAlphabetAutoAdd, "ReplaceTextFromSpellingChecker": local_ReplaceTextFromSpellingChecker, "OtherCorrectionsAutoAdd": local_OtherCorrectionsAutoAdd, "CorrectKeyboardSetting": local_CorrectKeyboardSetting, "CorrectTableCells": local_CorrectTableCells, "DisplayAutoCorrectOptions": local_DisplayAutoCorrectOptions, }


# Tool: 13
@mcp.tool()
async def word_get_FontNames():
	this_Global = EnsureWord()
	retVal = this_Global.FontNames
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "FontNames", "Count": local_Count, }


# Tool: 14
@mcp.tool()
async def word_get_LandscapeFontNames():
	this_Global = EnsureWord()
	retVal = this_Global.LandscapeFontNames
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "FontNames", "Count": local_Count, }


# Tool: 15
@mcp.tool()
async def word_get_PortraitFontNames():
	this_Global = EnsureWord()
	retVal = this_Global.PortraitFontNames
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "FontNames", "Count": local_Count, }


# Tool: 16
@mcp.tool()
async def word_get_Languages():
	this_Global = EnsureWord()
	retVal = this_Global.Languages
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Languages", "Count": local_Count, }


# Tool: 17
@mcp.tool()
async def word_get_Assistant():
	this_Global = EnsureWord()
	retVal = this_Global.Assistant
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Assistant"}


# Tool: 18
@mcp.tool()
async def word_get_FileConverters():
	this_Global = EnsureWord()
	retVal = this_Global.FileConverters
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	try:
		local_ConvertMacWordChevrons = retVal.ConvertMacWordChevrons
	except:
		local_ConvertMacWordChevrons = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "FileConverters", "Count": local_Count, "ConvertMacWordChevrons": local_ConvertMacWordChevrons, }


# Tool: 19
@mcp.tool()
async def word_get_Dialogs():
	this_Global = EnsureWord()
	retVal = this_Global.Dialogs
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Dialogs", "Count": local_Count, }


# Tool: 20
@mcp.tool()
async def word_get_CaptionLabels():
	this_Global = EnsureWord()
	retVal = this_Global.CaptionLabels
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "CaptionLabels", "Count": local_Count, }


# Tool: 21
@mcp.tool()
async def word_get_AutoCaptions():
	this_Global = EnsureWord()
	retVal = this_Global.AutoCaptions
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "AutoCaptions", "Count": local_Count, }


# Tool: 22
@mcp.tool()
async def word_get_AddIns():
	this_Global = EnsureWord()
	retVal = this_Global.AddIns
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "AddIns", "Count": local_Count, }


# Tool: 23
@mcp.tool()
async def word_get_Tasks():
	this_Global = EnsureWord()
	retVal = this_Global.Tasks
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Tasks", "Count": local_Count, }


# Tool: 24
@mcp.tool()
async def word_get_MacroContainer():
	this_Global = EnsureWord()
	retVal = this_Global.MacroContainer
	return retVal


# Tool: 25
@mcp.tool()
async def word_get_CommandBars():
	this_Global = EnsureWord()
	retVal = this_Global.CommandBars
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "CommandBars"}


# Tool: 26
@mcp.tool()
async def word_get_VBE():
	this_Global = EnsureWord()
	retVal = this_Global.VBE
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "VBE"}


# Tool: 27
@mcp.tool()
async def word_get_ListGalleries():
	this_Global = EnsureWord()
	retVal = this_Global.ListGalleries
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ListGalleries", "Count": local_Count, }


# Tool: 28
@mcp.tool()
async def word_get_ActivePrinter():
	this_Global = EnsureWord()
	retVal = this_Global.ActivePrinter
	return retVal


# Tool: 29
@mcp.tool()
async def word_get_Templates():
	this_Global = EnsureWord()
	retVal = this_Global.Templates
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Templates", "Count": local_Count, }


# Tool: 30
@mcp.tool()
async def word_get_CustomizationContext():
	this_Global = EnsureWord()
	retVal = this_Global.CustomizationContext
	return retVal


# Tool: 31
@mcp.tool()
async def word_get_KeyBindings():
	this_Global = EnsureWord()
	retVal = this_Global.KeyBindings
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	try:
		local_Context = retVal.Context
	except:
		local_Context = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "KeyBindings", "Count": local_Count, "Context": local_Context, }


# Tool: 32
@mcp.tool()
async def word_get_Options():
	this_Global = EnsureWord()
	retVal = this_Global.Options
	try:
		local_AllowAccentedUppercase = retVal.AllowAccentedUppercase
	except:
		local_AllowAccentedUppercase = None
	try:
		local_WPHelp = retVal.WPHelp
	except:
		local_WPHelp = None
	try:
		local_WPDocNavKeys = retVal.WPDocNavKeys
	except:
		local_WPDocNavKeys = None
	try:
		local_Pagination = retVal.Pagination
	except:
		local_Pagination = None
	try:
		local_BlueScreen = retVal.BlueScreen
	except:
		local_BlueScreen = None
	try:
		local_EnableSound = retVal.EnableSound
	except:
		local_EnableSound = None
	try:
		local_ConfirmConversions = retVal.ConfirmConversions
	except:
		local_ConfirmConversions = None
	try:
		local_UpdateLinksAtOpen = retVal.UpdateLinksAtOpen
	except:
		local_UpdateLinksAtOpen = None
	try:
		local_SendMailAttach = retVal.SendMailAttach
	except:
		local_SendMailAttach = None
	try:
		local_MeasurementUnit = retVal.MeasurementUnit
	except:
		local_MeasurementUnit = None
	try:
		local_ButtonFieldClicks = retVal.ButtonFieldClicks
	except:
		local_ButtonFieldClicks = None
	try:
		local_ShortMenuNames = retVal.ShortMenuNames
	except:
		local_ShortMenuNames = None
	try:
		local_RTFInClipboard = retVal.RTFInClipboard
	except:
		local_RTFInClipboard = None
	try:
		local_UpdateFieldsAtPrint = retVal.UpdateFieldsAtPrint
	except:
		local_UpdateFieldsAtPrint = None
	try:
		local_PrintProperties = retVal.PrintProperties
	except:
		local_PrintProperties = None
	try:
		local_PrintFieldCodes = retVal.PrintFieldCodes
	except:
		local_PrintFieldCodes = None
	try:
		local_PrintComments = retVal.PrintComments
	except:
		local_PrintComments = None
	try:
		local_PrintHiddenText = retVal.PrintHiddenText
	except:
		local_PrintHiddenText = None
	try:
		local_EnvelopeFeederInstalled = retVal.EnvelopeFeederInstalled
	except:
		local_EnvelopeFeederInstalled = None
	try:
		local_UpdateLinksAtPrint = retVal.UpdateLinksAtPrint
	except:
		local_UpdateLinksAtPrint = None
	try:
		local_PrintBackground = retVal.PrintBackground
	except:
		local_PrintBackground = None
	try:
		local_PrintDrawingObjects = retVal.PrintDrawingObjects
	except:
		local_PrintDrawingObjects = None
	try:
		local_DefaultTray = retVal.DefaultTray
	except:
		local_DefaultTray = None
	try:
		local_DefaultTrayID = retVal.DefaultTrayID
	except:
		local_DefaultTrayID = None
	try:
		local_CreateBackup = retVal.CreateBackup
	except:
		local_CreateBackup = None
	try:
		local_AllowFastSave = retVal.AllowFastSave
	except:
		local_AllowFastSave = None
	try:
		local_SavePropertiesPrompt = retVal.SavePropertiesPrompt
	except:
		local_SavePropertiesPrompt = None
	try:
		local_SaveNormalPrompt = retVal.SaveNormalPrompt
	except:
		local_SaveNormalPrompt = None
	try:
		local_SaveInterval = retVal.SaveInterval
	except:
		local_SaveInterval = None
	try:
		local_BackgroundSave = retVal.BackgroundSave
	except:
		local_BackgroundSave = None
	try:
		local_InsertedTextMark = retVal.InsertedTextMark
	except:
		local_InsertedTextMark = None
	try:
		local_DeletedTextMark = retVal.DeletedTextMark
	except:
		local_DeletedTextMark = None
	try:
		local_RevisedLinesMark = retVal.RevisedLinesMark
	except:
		local_RevisedLinesMark = None
	try:
		local_InsertedTextColor = retVal.InsertedTextColor
	except:
		local_InsertedTextColor = None
	try:
		local_DeletedTextColor = retVal.DeletedTextColor
	except:
		local_DeletedTextColor = None
	try:
		local_RevisedLinesColor = retVal.RevisedLinesColor
	except:
		local_RevisedLinesColor = None
	try:
		local_DefaultFilePath = retVal.DefaultFilePath
	except:
		local_DefaultFilePath = None
	try:
		local_Overtype = retVal.Overtype
	except:
		local_Overtype = None
	try:
		local_ReplaceSelection = retVal.ReplaceSelection
	except:
		local_ReplaceSelection = None
	try:
		local_AllowDragAndDrop = retVal.AllowDragAndDrop
	except:
		local_AllowDragAndDrop = None
	try:
		local_AutoWordSelection = retVal.AutoWordSelection
	except:
		local_AutoWordSelection = None
	try:
		local_INSKeyForPaste = retVal.INSKeyForPaste
	except:
		local_INSKeyForPaste = None
	try:
		local_SmartCutPaste = retVal.SmartCutPaste
	except:
		local_SmartCutPaste = None
	try:
		local_TabIndentKey = retVal.TabIndentKey
	except:
		local_TabIndentKey = None
	try:
		local_PictureEditor = retVal.PictureEditor
	except:
		local_PictureEditor = None
	try:
		local_AnimateScreenMovements = retVal.AnimateScreenMovements
	except:
		local_AnimateScreenMovements = None
	try:
		local_VirusProtection = retVal.VirusProtection
	except:
		local_VirusProtection = None
	try:
		local_RevisedPropertiesMark = retVal.RevisedPropertiesMark
	except:
		local_RevisedPropertiesMark = None
	try:
		local_RevisedPropertiesColor = retVal.RevisedPropertiesColor
	except:
		local_RevisedPropertiesColor = None
	try:
		local_SnapToGrid = retVal.SnapToGrid
	except:
		local_SnapToGrid = None
	try:
		local_SnapToShapes = retVal.SnapToShapes
	except:
		local_SnapToShapes = None
	try:
		local_GridDistanceHorizontal = retVal.GridDistanceHorizontal
	except:
		local_GridDistanceHorizontal = None
	try:
		local_GridDistanceVertical = retVal.GridDistanceVertical
	except:
		local_GridDistanceVertical = None
	try:
		local_GridOriginHorizontal = retVal.GridOriginHorizontal
	except:
		local_GridOriginHorizontal = None
	try:
		local_GridOriginVertical = retVal.GridOriginVertical
	except:
		local_GridOriginVertical = None
	try:
		local_InlineConversion = retVal.InlineConversion
	except:
		local_InlineConversion = None
	try:
		local_IMEAutomaticControl = retVal.IMEAutomaticControl
	except:
		local_IMEAutomaticControl = None
	try:
		local_AutoFormatApplyHeadings = retVal.AutoFormatApplyHeadings
	except:
		local_AutoFormatApplyHeadings = None
	try:
		local_AutoFormatApplyLists = retVal.AutoFormatApplyLists
	except:
		local_AutoFormatApplyLists = None
	try:
		local_AutoFormatApplyBulletedLists = retVal.AutoFormatApplyBulletedLists
	except:
		local_AutoFormatApplyBulletedLists = None
	try:
		local_AutoFormatApplyOtherParas = retVal.AutoFormatApplyOtherParas
	except:
		local_AutoFormatApplyOtherParas = None
	try:
		local_AutoFormatReplaceQuotes = retVal.AutoFormatReplaceQuotes
	except:
		local_AutoFormatReplaceQuotes = None
	try:
		local_AutoFormatReplaceSymbols = retVal.AutoFormatReplaceSymbols
	except:
		local_AutoFormatReplaceSymbols = None
	try:
		local_AutoFormatReplaceOrdinals = retVal.AutoFormatReplaceOrdinals
	except:
		local_AutoFormatReplaceOrdinals = None
	try:
		local_AutoFormatReplaceFractions = retVal.AutoFormatReplaceFractions
	except:
		local_AutoFormatReplaceFractions = None
	try:
		local_AutoFormatReplacePlainTextEmphasis = retVal.AutoFormatReplacePlainTextEmphasis
	except:
		local_AutoFormatReplacePlainTextEmphasis = None
	try:
		local_AutoFormatPreserveStyles = retVal.AutoFormatPreserveStyles
	except:
		local_AutoFormatPreserveStyles = None
	try:
		local_AutoFormatAsYouTypeApplyHeadings = retVal.AutoFormatAsYouTypeApplyHeadings
	except:
		local_AutoFormatAsYouTypeApplyHeadings = None
	try:
		local_AutoFormatAsYouTypeApplyBorders = retVal.AutoFormatAsYouTypeApplyBorders
	except:
		local_AutoFormatAsYouTypeApplyBorders = None
	try:
		local_AutoFormatAsYouTypeApplyBulletedLists = retVal.AutoFormatAsYouTypeApplyBulletedLists
	except:
		local_AutoFormatAsYouTypeApplyBulletedLists = None
	try:
		local_AutoFormatAsYouTypeApplyNumberedLists = retVal.AutoFormatAsYouTypeApplyNumberedLists
	except:
		local_AutoFormatAsYouTypeApplyNumberedLists = None
	try:
		local_AutoFormatAsYouTypeReplaceQuotes = retVal.AutoFormatAsYouTypeReplaceQuotes
	except:
		local_AutoFormatAsYouTypeReplaceQuotes = None
	try:
		local_AutoFormatAsYouTypeReplaceSymbols = retVal.AutoFormatAsYouTypeReplaceSymbols
	except:
		local_AutoFormatAsYouTypeReplaceSymbols = None
	try:
		local_AutoFormatAsYouTypeReplaceOrdinals = retVal.AutoFormatAsYouTypeReplaceOrdinals
	except:
		local_AutoFormatAsYouTypeReplaceOrdinals = None
	try:
		local_AutoFormatAsYouTypeReplaceFractions = retVal.AutoFormatAsYouTypeReplaceFractions
	except:
		local_AutoFormatAsYouTypeReplaceFractions = None
	try:
		local_AutoFormatAsYouTypeReplacePlainTextEmphasis = retVal.AutoFormatAsYouTypeReplacePlainTextEmphasis
	except:
		local_AutoFormatAsYouTypeReplacePlainTextEmphasis = None
	try:
		local_AutoFormatAsYouTypeFormatListItemBeginning = retVal.AutoFormatAsYouTypeFormatListItemBeginning
	except:
		local_AutoFormatAsYouTypeFormatListItemBeginning = None
	try:
		local_AutoFormatAsYouTypeDefineStyles = retVal.AutoFormatAsYouTypeDefineStyles
	except:
		local_AutoFormatAsYouTypeDefineStyles = None
	try:
		local_AutoFormatPlainTextWordMail = retVal.AutoFormatPlainTextWordMail
	except:
		local_AutoFormatPlainTextWordMail = None
	try:
		local_AutoFormatAsYouTypeReplaceHyperlinks = retVal.AutoFormatAsYouTypeReplaceHyperlinks
	except:
		local_AutoFormatAsYouTypeReplaceHyperlinks = None
	try:
		local_AutoFormatReplaceHyperlinks = retVal.AutoFormatReplaceHyperlinks
	except:
		local_AutoFormatReplaceHyperlinks = None
	try:
		local_DefaultHighlightColorIndex = retVal.DefaultHighlightColorIndex
	except:
		local_DefaultHighlightColorIndex = None
	try:
		local_DefaultBorderLineStyle = retVal.DefaultBorderLineStyle
	except:
		local_DefaultBorderLineStyle = None
	try:
		local_CheckSpellingAsYouType = retVal.CheckSpellingAsYouType
	except:
		local_CheckSpellingAsYouType = None
	try:
		local_CheckGrammarAsYouType = retVal.CheckGrammarAsYouType
	except:
		local_CheckGrammarAsYouType = None
	try:
		local_IgnoreInternetAndFileAddresses = retVal.IgnoreInternetAndFileAddresses
	except:
		local_IgnoreInternetAndFileAddresses = None
	try:
		local_ShowReadabilityStatistics = retVal.ShowReadabilityStatistics
	except:
		local_ShowReadabilityStatistics = None
	try:
		local_IgnoreUppercase = retVal.IgnoreUppercase
	except:
		local_IgnoreUppercase = None
	try:
		local_IgnoreMixedDigits = retVal.IgnoreMixedDigits
	except:
		local_IgnoreMixedDigits = None
	try:
		local_SuggestFromMainDictionaryOnly = retVal.SuggestFromMainDictionaryOnly
	except:
		local_SuggestFromMainDictionaryOnly = None
	try:
		local_SuggestSpellingCorrections = retVal.SuggestSpellingCorrections
	except:
		local_SuggestSpellingCorrections = None
	try:
		local_DefaultBorderLineWidth = retVal.DefaultBorderLineWidth
	except:
		local_DefaultBorderLineWidth = None
	try:
		local_CheckGrammarWithSpelling = retVal.CheckGrammarWithSpelling
	except:
		local_CheckGrammarWithSpelling = None
	try:
		local_DefaultOpenFormat = retVal.DefaultOpenFormat
	except:
		local_DefaultOpenFormat = None
	try:
		local_PrintDraft = retVal.PrintDraft
	except:
		local_PrintDraft = None
	try:
		local_PrintReverse = retVal.PrintReverse
	except:
		local_PrintReverse = None
	try:
		local_MapPaperSize = retVal.MapPaperSize
	except:
		local_MapPaperSize = None
	try:
		local_AutoFormatAsYouTypeApplyTables = retVal.AutoFormatAsYouTypeApplyTables
	except:
		local_AutoFormatAsYouTypeApplyTables = None
	try:
		local_AutoFormatApplyFirstIndents = retVal.AutoFormatApplyFirstIndents
	except:
		local_AutoFormatApplyFirstIndents = None
	try:
		local_AutoFormatMatchParentheses = retVal.AutoFormatMatchParentheses
	except:
		local_AutoFormatMatchParentheses = None
	try:
		local_AutoFormatReplaceFarEastDashes = retVal.AutoFormatReplaceFarEastDashes
	except:
		local_AutoFormatReplaceFarEastDashes = None
	try:
		local_AutoFormatDeleteAutoSpaces = retVal.AutoFormatDeleteAutoSpaces
	except:
		local_AutoFormatDeleteAutoSpaces = None
	try:
		local_AutoFormatAsYouTypeApplyFirstIndents = retVal.AutoFormatAsYouTypeApplyFirstIndents
	except:
		local_AutoFormatAsYouTypeApplyFirstIndents = None
	try:
		local_AutoFormatAsYouTypeApplyDates = retVal.AutoFormatAsYouTypeApplyDates
	except:
		local_AutoFormatAsYouTypeApplyDates = None
	try:
		local_AutoFormatAsYouTypeApplyClosings = retVal.AutoFormatAsYouTypeApplyClosings
	except:
		local_AutoFormatAsYouTypeApplyClosings = None
	try:
		local_AutoFormatAsYouTypeMatchParentheses = retVal.AutoFormatAsYouTypeMatchParentheses
	except:
		local_AutoFormatAsYouTypeMatchParentheses = None
	try:
		local_AutoFormatAsYouTypeReplaceFarEastDashes = retVal.AutoFormatAsYouTypeReplaceFarEastDashes
	except:
		local_AutoFormatAsYouTypeReplaceFarEastDashes = None
	try:
		local_AutoFormatAsYouTypeDeleteAutoSpaces = retVal.AutoFormatAsYouTypeDeleteAutoSpaces
	except:
		local_AutoFormatAsYouTypeDeleteAutoSpaces = None
	try:
		local_AutoFormatAsYouTypeInsertClosings = retVal.AutoFormatAsYouTypeInsertClosings
	except:
		local_AutoFormatAsYouTypeInsertClosings = None
	try:
		local_AutoFormatAsYouTypeAutoLetterWizard = retVal.AutoFormatAsYouTypeAutoLetterWizard
	except:
		local_AutoFormatAsYouTypeAutoLetterWizard = None
	try:
		local_AutoFormatAsYouTypeInsertOvers = retVal.AutoFormatAsYouTypeInsertOvers
	except:
		local_AutoFormatAsYouTypeInsertOvers = None
	try:
		local_DisplayGridLines = retVal.DisplayGridLines
	except:
		local_DisplayGridLines = None
	try:
		local_MatchFuzzyCase = retVal.MatchFuzzyCase
	except:
		local_MatchFuzzyCase = None
	try:
		local_MatchFuzzyByte = retVal.MatchFuzzyByte
	except:
		local_MatchFuzzyByte = None
	try:
		local_MatchFuzzyHiragana = retVal.MatchFuzzyHiragana
	except:
		local_MatchFuzzyHiragana = None
	try:
		local_MatchFuzzySmallKana = retVal.MatchFuzzySmallKana
	except:
		local_MatchFuzzySmallKana = None
	try:
		local_MatchFuzzyDash = retVal.MatchFuzzyDash
	except:
		local_MatchFuzzyDash = None
	try:
		local_MatchFuzzyIterationMark = retVal.MatchFuzzyIterationMark
	except:
		local_MatchFuzzyIterationMark = None
	try:
		local_MatchFuzzyKanji = retVal.MatchFuzzyKanji
	except:
		local_MatchFuzzyKanji = None
	try:
		local_MatchFuzzyOldKana = retVal.MatchFuzzyOldKana
	except:
		local_MatchFuzzyOldKana = None
	try:
		local_MatchFuzzyProlongedSoundMark = retVal.MatchFuzzyProlongedSoundMark
	except:
		local_MatchFuzzyProlongedSoundMark = None
	try:
		local_MatchFuzzyDZ = retVal.MatchFuzzyDZ
	except:
		local_MatchFuzzyDZ = None
	try:
		local_MatchFuzzyBV = retVal.MatchFuzzyBV
	except:
		local_MatchFuzzyBV = None
	try:
		local_MatchFuzzyTC = retVal.MatchFuzzyTC
	except:
		local_MatchFuzzyTC = None
	try:
		local_MatchFuzzyHF = retVal.MatchFuzzyHF
	except:
		local_MatchFuzzyHF = None
	try:
		local_MatchFuzzyZJ = retVal.MatchFuzzyZJ
	except:
		local_MatchFuzzyZJ = None
	try:
		local_MatchFuzzyAY = retVal.MatchFuzzyAY
	except:
		local_MatchFuzzyAY = None
	try:
		local_MatchFuzzyKiKu = retVal.MatchFuzzyKiKu
	except:
		local_MatchFuzzyKiKu = None
	try:
		local_MatchFuzzyPunctuation = retVal.MatchFuzzyPunctuation
	except:
		local_MatchFuzzyPunctuation = None
	try:
		local_MatchFuzzySpace = retVal.MatchFuzzySpace
	except:
		local_MatchFuzzySpace = None
	try:
		local_ApplyFarEastFontsToAscii = retVal.ApplyFarEastFontsToAscii
	except:
		local_ApplyFarEastFontsToAscii = None
	try:
		local_ConvertHighAnsiToFarEast = retVal.ConvertHighAnsiToFarEast
	except:
		local_ConvertHighAnsiToFarEast = None
	try:
		local_PrintOddPagesInAscendingOrder = retVal.PrintOddPagesInAscendingOrder
	except:
		local_PrintOddPagesInAscendingOrder = None
	try:
		local_PrintEvenPagesInAscendingOrder = retVal.PrintEvenPagesInAscendingOrder
	except:
		local_PrintEvenPagesInAscendingOrder = None
	try:
		local_DefaultBorderColorIndex = retVal.DefaultBorderColorIndex
	except:
		local_DefaultBorderColorIndex = None
	try:
		local_EnableMisusedWordsDictionary = retVal.EnableMisusedWordsDictionary
	except:
		local_EnableMisusedWordsDictionary = None
	try:
		local_AllowCombinedAuxiliaryForms = retVal.AllowCombinedAuxiliaryForms
	except:
		local_AllowCombinedAuxiliaryForms = None
	try:
		local_HangulHanjaFastConversion = retVal.HangulHanjaFastConversion
	except:
		local_HangulHanjaFastConversion = None
	try:
		local_CheckHangulEndings = retVal.CheckHangulEndings
	except:
		local_CheckHangulEndings = None
	try:
		local_EnableHangulHanjaRecentOrdering = retVal.EnableHangulHanjaRecentOrdering
	except:
		local_EnableHangulHanjaRecentOrdering = None
	try:
		local_MultipleWordConversionsMode = retVal.MultipleWordConversionsMode
	except:
		local_MultipleWordConversionsMode = None
	try:
		local_DefaultBorderColor = retVal.DefaultBorderColor
	except:
		local_DefaultBorderColor = None
	try:
		local_AllowPixelUnits = retVal.AllowPixelUnits
	except:
		local_AllowPixelUnits = None
	try:
		local_UseCharacterUnit = retVal.UseCharacterUnit
	except:
		local_UseCharacterUnit = None
	try:
		local_AllowCompoundNounProcessing = retVal.AllowCompoundNounProcessing
	except:
		local_AllowCompoundNounProcessing = None
	try:
		local_AutoKeyboardSwitching = retVal.AutoKeyboardSwitching
	except:
		local_AutoKeyboardSwitching = None
	try:
		local_DocumentViewDirection = retVal.DocumentViewDirection
	except:
		local_DocumentViewDirection = None
	try:
		local_ArabicNumeral = retVal.ArabicNumeral
	except:
		local_ArabicNumeral = None
	try:
		local_MonthNames = retVal.MonthNames
	except:
		local_MonthNames = None
	try:
		local_CursorMovement = retVal.CursorMovement
	except:
		local_CursorMovement = None
	try:
		local_VisualSelection = retVal.VisualSelection
	except:
		local_VisualSelection = None
	try:
		local_ShowDiacritics = retVal.ShowDiacritics
	except:
		local_ShowDiacritics = None
	try:
		local_ShowControlCharacters = retVal.ShowControlCharacters
	except:
		local_ShowControlCharacters = None
	try:
		local_AddControlCharacters = retVal.AddControlCharacters
	except:
		local_AddControlCharacters = None
	try:
		local_AddBiDirectionalMarksWhenSavingTextFile = retVal.AddBiDirectionalMarksWhenSavingTextFile
	except:
		local_AddBiDirectionalMarksWhenSavingTextFile = None
	try:
		local_StrictInitialAlefHamza = retVal.StrictInitialAlefHamza
	except:
		local_StrictInitialAlefHamza = None
	try:
		local_StrictFinalYaa = retVal.StrictFinalYaa
	except:
		local_StrictFinalYaa = None
	try:
		local_HebrewMode = retVal.HebrewMode
	except:
		local_HebrewMode = None
	try:
		local_ArabicMode = retVal.ArabicMode
	except:
		local_ArabicMode = None
	try:
		local_AllowClickAndTypeMouse = retVal.AllowClickAndTypeMouse
	except:
		local_AllowClickAndTypeMouse = None
	try:
		local_UseGermanSpellingReform = retVal.UseGermanSpellingReform
	except:
		local_UseGermanSpellingReform = None
	try:
		local_InterpretHighAnsi = retVal.InterpretHighAnsi
	except:
		local_InterpretHighAnsi = None
	try:
		local_AddHebDoubleQuote = retVal.AddHebDoubleQuote
	except:
		local_AddHebDoubleQuote = None
	try:
		local_UseDiffDiacColor = retVal.UseDiffDiacColor
	except:
		local_UseDiffDiacColor = None
	try:
		local_DiacriticColorVal = retVal.DiacriticColorVal
	except:
		local_DiacriticColorVal = None
	try:
		local_OptimizeForWord97byDefault = retVal.OptimizeForWord97byDefault
	except:
		local_OptimizeForWord97byDefault = None
	try:
		local_LocalNetworkFile = retVal.LocalNetworkFile
	except:
		local_LocalNetworkFile = None
	try:
		local_TypeNReplace = retVal.TypeNReplace
	except:
		local_TypeNReplace = None
	try:
		local_SequenceCheck = retVal.SequenceCheck
	except:
		local_SequenceCheck = None
	try:
		local_BackgroundOpen = retVal.BackgroundOpen
	except:
		local_BackgroundOpen = None
	try:
		local_DisableFeaturesbyDefault = retVal.DisableFeaturesbyDefault
	except:
		local_DisableFeaturesbyDefault = None
	try:
		local_PasteAdjustWordSpacing = retVal.PasteAdjustWordSpacing
	except:
		local_PasteAdjustWordSpacing = None
	try:
		local_PasteAdjustParagraphSpacing = retVal.PasteAdjustParagraphSpacing
	except:
		local_PasteAdjustParagraphSpacing = None
	try:
		local_PasteAdjustTableFormatting = retVal.PasteAdjustTableFormatting
	except:
		local_PasteAdjustTableFormatting = None
	try:
		local_PasteSmartStyleBehavior = retVal.PasteSmartStyleBehavior
	except:
		local_PasteSmartStyleBehavior = None
	try:
		local_PasteMergeFromPPT = retVal.PasteMergeFromPPT
	except:
		local_PasteMergeFromPPT = None
	try:
		local_PasteMergeFromXL = retVal.PasteMergeFromXL
	except:
		local_PasteMergeFromXL = None
	try:
		local_CtrlClickHyperlinkToOpen = retVal.CtrlClickHyperlinkToOpen
	except:
		local_CtrlClickHyperlinkToOpen = None
	try:
		local_PictureWrapType = retVal.PictureWrapType
	except:
		local_PictureWrapType = None
	try:
		local_DisableFeaturesIntroducedAfterbyDefault = retVal.DisableFeaturesIntroducedAfterbyDefault
	except:
		local_DisableFeaturesIntroducedAfterbyDefault = None
	try:
		local_PasteSmartCutPaste = retVal.PasteSmartCutPaste
	except:
		local_PasteSmartCutPaste = None
	try:
		local_DisplayPasteOptions = retVal.DisplayPasteOptions
	except:
		local_DisplayPasteOptions = None
	try:
		local_PromptUpdateStyle = retVal.PromptUpdateStyle
	except:
		local_PromptUpdateStyle = None
	try:
		local_DefaultEPostageApp = retVal.DefaultEPostageApp
	except:
		local_DefaultEPostageApp = None
	try:
		local_DefaultTextEncoding = retVal.DefaultTextEncoding
	except:
		local_DefaultTextEncoding = None
	try:
		local_LabelSmartTags = retVal.LabelSmartTags
	except:
		local_LabelSmartTags = None
	try:
		local_DisplaySmartTagButtons = retVal.DisplaySmartTagButtons
	except:
		local_DisplaySmartTagButtons = None
	try:
		local_WarnBeforeSavingPrintingSendingMarkup = retVal.WarnBeforeSavingPrintingSendingMarkup
	except:
		local_WarnBeforeSavingPrintingSendingMarkup = None
	try:
		local_StoreRSIDOnSave = retVal.StoreRSIDOnSave
	except:
		local_StoreRSIDOnSave = None
	try:
		local_ShowFormatError = retVal.ShowFormatError
	except:
		local_ShowFormatError = None
	try:
		local_FormatScanning = retVal.FormatScanning
	except:
		local_FormatScanning = None
	try:
		local_PasteMergeLists = retVal.PasteMergeLists
	except:
		local_PasteMergeLists = None
	try:
		local_AutoCreateNewDrawings = retVal.AutoCreateNewDrawings
	except:
		local_AutoCreateNewDrawings = None
	try:
		local_SmartParaSelection = retVal.SmartParaSelection
	except:
		local_SmartParaSelection = None
	try:
		local_RevisionsBalloonPrintOrientation = retVal.RevisionsBalloonPrintOrientation
	except:
		local_RevisionsBalloonPrintOrientation = None
	try:
		local_CommentsColor = retVal.CommentsColor
	except:
		local_CommentsColor = None
	try:
		local_PrintXMLTag = retVal.PrintXMLTag
	except:
		local_PrintXMLTag = None
	try:
		local_PrintBackgrounds = retVal.PrintBackgrounds
	except:
		local_PrintBackgrounds = None
	try:
		local_AllowReadingMode = retVal.AllowReadingMode
	except:
		local_AllowReadingMode = None
	try:
		local_ShowMarkupOpenSave = retVal.ShowMarkupOpenSave
	except:
		local_ShowMarkupOpenSave = None
	try:
		local_SmartCursoring = retVal.SmartCursoring
	except:
		local_SmartCursoring = None
	try:
		local_MoveToTextMark = retVal.MoveToTextMark
	except:
		local_MoveToTextMark = None
	try:
		local_MoveFromTextMark = retVal.MoveFromTextMark
	except:
		local_MoveFromTextMark = None
	try:
		local_BibliographyStyle = retVal.BibliographyStyle
	except:
		local_BibliographyStyle = None
	try:
		local_BibliographySort = retVal.BibliographySort
	except:
		local_BibliographySort = None
	try:
		local_InsertedCellColor = retVal.InsertedCellColor
	except:
		local_InsertedCellColor = None
	try:
		local_DeletedCellColor = retVal.DeletedCellColor
	except:
		local_DeletedCellColor = None
	try:
		local_MergedCellColor = retVal.MergedCellColor
	except:
		local_MergedCellColor = None
	try:
		local_SplitCellColor = retVal.SplitCellColor
	except:
		local_SplitCellColor = None
	try:
		local_ShowSelectionFloaties = retVal.ShowSelectionFloaties
	except:
		local_ShowSelectionFloaties = None
	try:
		local_ShowMenuFloaties = retVal.ShowMenuFloaties
	except:
		local_ShowMenuFloaties = None
	try:
		local_ShowDevTools = retVal.ShowDevTools
	except:
		local_ShowDevTools = None
	try:
		local_EnableLivePreview = retVal.EnableLivePreview
	except:
		local_EnableLivePreview = None
	try:
		local_OMathAutoBuildUp = retVal.OMathAutoBuildUp
	except:
		local_OMathAutoBuildUp = None
	try:
		local_AlwaysUseClearType = retVal.AlwaysUseClearType
	except:
		local_AlwaysUseClearType = None
	try:
		local_PasteFormatWithinDocument = retVal.PasteFormatWithinDocument
	except:
		local_PasteFormatWithinDocument = None
	try:
		local_PasteFormatBetweenDocuments = retVal.PasteFormatBetweenDocuments
	except:
		local_PasteFormatBetweenDocuments = None
	try:
		local_PasteFormatBetweenStyledDocuments = retVal.PasteFormatBetweenStyledDocuments
	except:
		local_PasteFormatBetweenStyledDocuments = None
	try:
		local_PasteFormatFromExternalSource = retVal.PasteFormatFromExternalSource
	except:
		local_PasteFormatFromExternalSource = None
	try:
		local_PasteOptionKeepBulletsAndNumbers = retVal.PasteOptionKeepBulletsAndNumbers
	except:
		local_PasteOptionKeepBulletsAndNumbers = None
	try:
		local_INSKeyForOvertype = retVal.INSKeyForOvertype
	except:
		local_INSKeyForOvertype = None
	try:
		local_RepeatWord = retVal.RepeatWord
	except:
		local_RepeatWord = None
	try:
		local_FrenchReform = retVal.FrenchReform
	except:
		local_FrenchReform = None
	try:
		local_ContextualSpeller = retVal.ContextualSpeller
	except:
		local_ContextualSpeller = None
	try:
		local_MoveToTextColor = retVal.MoveToTextColor
	except:
		local_MoveToTextColor = None
	try:
		local_MoveFromTextColor = retVal.MoveFromTextColor
	except:
		local_MoveFromTextColor = None
	try:
		local_OMathCopyLF = retVal.OMathCopyLF
	except:
		local_OMathCopyLF = None
	try:
		local_UseNormalStyleForList = retVal.UseNormalStyleForList
	except:
		local_UseNormalStyleForList = None
	try:
		local_AllowOpenInDraftView = retVal.AllowOpenInDraftView
	except:
		local_AllowOpenInDraftView = None
	try:
		local_EnableLegacyIMEMode = retVal.EnableLegacyIMEMode
	except:
		local_EnableLegacyIMEMode = None
	try:
		local_DoNotPromptForConvert = retVal.DoNotPromptForConvert
	except:
		local_DoNotPromptForConvert = None
	try:
		local_PrecisePositioning = retVal.PrecisePositioning
	except:
		local_PrecisePositioning = None
	try:
		local_UpdateStyleListBehavior = retVal.UpdateStyleListBehavior
	except:
		local_UpdateStyleListBehavior = None
	try:
		local_StrictTaaMarboota = retVal.StrictTaaMarboota
	except:
		local_StrictTaaMarboota = None
	try:
		local_StrictRussianE = retVal.StrictRussianE
	except:
		local_StrictRussianE = None
	try:
		local_SpanishMode = retVal.SpanishMode
	except:
		local_SpanishMode = None
	try:
		local_PortugalReform = retVal.PortugalReform
	except:
		local_PortugalReform = None
	try:
		local_BrazilReform = retVal.BrazilReform
	except:
		local_BrazilReform = None
	try:
		local_UpdateFieldsWithTrackedChangesAtPrint = retVal.UpdateFieldsWithTrackedChangesAtPrint
	except:
		local_UpdateFieldsWithTrackedChangesAtPrint = None
	try:
		local_DisplayAlignmentGuides = retVal.DisplayAlignmentGuides
	except:
		local_DisplayAlignmentGuides = None
	try:
		local_PageAlignmentGuides = retVal.PageAlignmentGuides
	except:
		local_PageAlignmentGuides = None
	try:
		local_MarginAlignmentGuides = retVal.MarginAlignmentGuides
	except:
		local_MarginAlignmentGuides = None
	try:
		local_ParagraphAlignmentGuides = retVal.ParagraphAlignmentGuides
	except:
		local_ParagraphAlignmentGuides = None
	try:
		local_EnableLiveDrag = retVal.EnableLiveDrag
	except:
		local_EnableLiveDrag = None
	try:
		local_UseSubPixelPositioning = retVal.UseSubPixelPositioning
	except:
		local_UseSubPixelPositioning = None
	try:
		local_AlertIfNotDefault = retVal.AlertIfNotDefault
	except:
		local_AlertIfNotDefault = None
	try:
		local_EnableProofingToolsAdvertisement = retVal.EnableProofingToolsAdvertisement
	except:
		local_EnableProofingToolsAdvertisement = None
	try:
		local_PreferCloudSaveLocations = retVal.PreferCloudSaveLocations
	except:
		local_PreferCloudSaveLocations = None
	try:
		local_SkyDriveSignInOption = retVal.SkyDriveSignInOption
	except:
		local_SkyDriveSignInOption = None
	try:
		local_ExpandHeadingsOnOpen = retVal.ExpandHeadingsOnOpen
	except:
		local_ExpandHeadingsOnOpen = None
	try:
		local_UseLocalUserInfo = retVal.UseLocalUserInfo
	except:
		local_UseLocalUserInfo = None
	try:
		local_CloudSignInOption = retVal.CloudSignInOption
	except:
		local_CloudSignInOption = None
	try:
		local_ShowPopupAddRowColToTable = retVal.ShowPopupAddRowColToTable
	except:
		local_ShowPopupAddRowColToTable = None
	try:
		local_LiveWordCount = retVal.LiveWordCount
	except:
		local_LiveWordCount = None
	try:
		local_AllowCoAuthoringOnFilesWithMacros = retVal.AllowCoAuthoringOnFilesWithMacros
	except:
		local_AllowCoAuthoringOnFilesWithMacros = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Options", "AllowAccentedUppercase": local_AllowAccentedUppercase, "WPHelp": local_WPHelp, "WPDocNavKeys": local_WPDocNavKeys, "Pagination": local_Pagination, "BlueScreen": local_BlueScreen, "EnableSound": local_EnableSound, "ConfirmConversions": local_ConfirmConversions, "UpdateLinksAtOpen": local_UpdateLinksAtOpen, "SendMailAttach": local_SendMailAttach, "MeasurementUnit": local_MeasurementUnit, "ButtonFieldClicks": local_ButtonFieldClicks, "ShortMenuNames": local_ShortMenuNames, "RTFInClipboard": local_RTFInClipboard, "UpdateFieldsAtPrint": local_UpdateFieldsAtPrint, "PrintProperties": local_PrintProperties, "PrintFieldCodes": local_PrintFieldCodes, "PrintComments": local_PrintComments, "PrintHiddenText": local_PrintHiddenText, "EnvelopeFeederInstalled": local_EnvelopeFeederInstalled, "UpdateLinksAtPrint": local_UpdateLinksAtPrint, "PrintBackground": local_PrintBackground, "PrintDrawingObjects": local_PrintDrawingObjects, "DefaultTray": local_DefaultTray, "DefaultTrayID": local_DefaultTrayID, "CreateBackup": local_CreateBackup, "AllowFastSave": local_AllowFastSave, "SavePropertiesPrompt": local_SavePropertiesPrompt, "SaveNormalPrompt": local_SaveNormalPrompt, "SaveInterval": local_SaveInterval, "BackgroundSave": local_BackgroundSave, "InsertedTextMark": local_InsertedTextMark, "DeletedTextMark": local_DeletedTextMark, "RevisedLinesMark": local_RevisedLinesMark, "InsertedTextColor": local_InsertedTextColor, "DeletedTextColor": local_DeletedTextColor, "RevisedLinesColor": local_RevisedLinesColor, "DefaultFilePath": local_DefaultFilePath, "Overtype": local_Overtype, "ReplaceSelection": local_ReplaceSelection, "AllowDragAndDrop": local_AllowDragAndDrop, "AutoWordSelection": local_AutoWordSelection, "INSKeyForPaste": local_INSKeyForPaste, "SmartCutPaste": local_SmartCutPaste, "TabIndentKey": local_TabIndentKey, "PictureEditor": local_PictureEditor, "AnimateScreenMovements": local_AnimateScreenMovements, "VirusProtection": local_VirusProtection, "RevisedPropertiesMark": local_RevisedPropertiesMark, "RevisedPropertiesColor": local_RevisedPropertiesColor, "SnapToGrid": local_SnapToGrid, "SnapToShapes": local_SnapToShapes, "GridDistanceHorizontal": local_GridDistanceHorizontal, "GridDistanceVertical": local_GridDistanceVertical, "GridOriginHorizontal": local_GridOriginHorizontal, "GridOriginVertical": local_GridOriginVertical, "InlineConversion": local_InlineConversion, "IMEAutomaticControl": local_IMEAutomaticControl, "AutoFormatApplyHeadings": local_AutoFormatApplyHeadings, "AutoFormatApplyLists": local_AutoFormatApplyLists, "AutoFormatApplyBulletedLists": local_AutoFormatApplyBulletedLists, "AutoFormatApplyOtherParas": local_AutoFormatApplyOtherParas, "AutoFormatReplaceQuotes": local_AutoFormatReplaceQuotes, "AutoFormatReplaceSymbols": local_AutoFormatReplaceSymbols, "AutoFormatReplaceOrdinals": local_AutoFormatReplaceOrdinals, "AutoFormatReplaceFractions": local_AutoFormatReplaceFractions, "AutoFormatReplacePlainTextEmphasis": local_AutoFormatReplacePlainTextEmphasis, "AutoFormatPreserveStyles": local_AutoFormatPreserveStyles, "AutoFormatAsYouTypeApplyHeadings": local_AutoFormatAsYouTypeApplyHeadings, "AutoFormatAsYouTypeApplyBorders": local_AutoFormatAsYouTypeApplyBorders, "AutoFormatAsYouTypeApplyBulletedLists": local_AutoFormatAsYouTypeApplyBulletedLists, "AutoFormatAsYouTypeApplyNumberedLists": local_AutoFormatAsYouTypeApplyNumberedLists, "AutoFormatAsYouTypeReplaceQuotes": local_AutoFormatAsYouTypeReplaceQuotes, "AutoFormatAsYouTypeReplaceSymbols": local_AutoFormatAsYouTypeReplaceSymbols, "AutoFormatAsYouTypeReplaceOrdinals": local_AutoFormatAsYouTypeReplaceOrdinals, "AutoFormatAsYouTypeReplaceFractions": local_AutoFormatAsYouTypeReplaceFractions, "AutoFormatAsYouTypeReplacePlainTextEmphasis": local_AutoFormatAsYouTypeReplacePlainTextEmphasis, "AutoFormatAsYouTypeFormatListItemBeginning": local_AutoFormatAsYouTypeFormatListItemBeginning, "AutoFormatAsYouTypeDefineStyles": local_AutoFormatAsYouTypeDefineStyles, "AutoFormatPlainTextWordMail": local_AutoFormatPlainTextWordMail, "AutoFormatAsYouTypeReplaceHyperlinks": local_AutoFormatAsYouTypeReplaceHyperlinks, "AutoFormatReplaceHyperlinks": local_AutoFormatReplaceHyperlinks, "DefaultHighlightColorIndex": local_DefaultHighlightColorIndex, "DefaultBorderLineStyle": local_DefaultBorderLineStyle, "CheckSpellingAsYouType": local_CheckSpellingAsYouType, "CheckGrammarAsYouType": local_CheckGrammarAsYouType, "IgnoreInternetAndFileAddresses": local_IgnoreInternetAndFileAddresses, "ShowReadabilityStatistics": local_ShowReadabilityStatistics, "IgnoreUppercase": local_IgnoreUppercase, "IgnoreMixedDigits": local_IgnoreMixedDigits, "SuggestFromMainDictionaryOnly": local_SuggestFromMainDictionaryOnly, "SuggestSpellingCorrections": local_SuggestSpellingCorrections, "DefaultBorderLineWidth": local_DefaultBorderLineWidth, "CheckGrammarWithSpelling": local_CheckGrammarWithSpelling, "DefaultOpenFormat": local_DefaultOpenFormat, "PrintDraft": local_PrintDraft, "PrintReverse": local_PrintReverse, "MapPaperSize": local_MapPaperSize, "AutoFormatAsYouTypeApplyTables": local_AutoFormatAsYouTypeApplyTables, "AutoFormatApplyFirstIndents": local_AutoFormatApplyFirstIndents, "AutoFormatMatchParentheses": local_AutoFormatMatchParentheses, "AutoFormatReplaceFarEastDashes": local_AutoFormatReplaceFarEastDashes, "AutoFormatDeleteAutoSpaces": local_AutoFormatDeleteAutoSpaces, "AutoFormatAsYouTypeApplyFirstIndents": local_AutoFormatAsYouTypeApplyFirstIndents, "AutoFormatAsYouTypeApplyDates": local_AutoFormatAsYouTypeApplyDates, "AutoFormatAsYouTypeApplyClosings": local_AutoFormatAsYouTypeApplyClosings, "AutoFormatAsYouTypeMatchParentheses": local_AutoFormatAsYouTypeMatchParentheses, "AutoFormatAsYouTypeReplaceFarEastDashes": local_AutoFormatAsYouTypeReplaceFarEastDashes, "AutoFormatAsYouTypeDeleteAutoSpaces": local_AutoFormatAsYouTypeDeleteAutoSpaces, "AutoFormatAsYouTypeInsertClosings": local_AutoFormatAsYouTypeInsertClosings, "AutoFormatAsYouTypeAutoLetterWizard": local_AutoFormatAsYouTypeAutoLetterWizard, "AutoFormatAsYouTypeInsertOvers": local_AutoFormatAsYouTypeInsertOvers, "DisplayGridLines": local_DisplayGridLines, "MatchFuzzyCase": local_MatchFuzzyCase, "MatchFuzzyByte": local_MatchFuzzyByte, "MatchFuzzyHiragana": local_MatchFuzzyHiragana, "MatchFuzzySmallKana": local_MatchFuzzySmallKana, "MatchFuzzyDash": local_MatchFuzzyDash, "MatchFuzzyIterationMark": local_MatchFuzzyIterationMark, "MatchFuzzyKanji": local_MatchFuzzyKanji, "MatchFuzzyOldKana": local_MatchFuzzyOldKana, "MatchFuzzyProlongedSoundMark": local_MatchFuzzyProlongedSoundMark, "MatchFuzzyDZ": local_MatchFuzzyDZ, "MatchFuzzyBV": local_MatchFuzzyBV, "MatchFuzzyTC": local_MatchFuzzyTC, "MatchFuzzyHF": local_MatchFuzzyHF, "MatchFuzzyZJ": local_MatchFuzzyZJ, "MatchFuzzyAY": local_MatchFuzzyAY, "MatchFuzzyKiKu": local_MatchFuzzyKiKu, "MatchFuzzyPunctuation": local_MatchFuzzyPunctuation, "MatchFuzzySpace": local_MatchFuzzySpace, "ApplyFarEastFontsToAscii": local_ApplyFarEastFontsToAscii, "ConvertHighAnsiToFarEast": local_ConvertHighAnsiToFarEast, "PrintOddPagesInAscendingOrder": local_PrintOddPagesInAscendingOrder, "PrintEvenPagesInAscendingOrder": local_PrintEvenPagesInAscendingOrder, "DefaultBorderColorIndex": local_DefaultBorderColorIndex, "EnableMisusedWordsDictionary": local_EnableMisusedWordsDictionary, "AllowCombinedAuxiliaryForms": local_AllowCombinedAuxiliaryForms, "HangulHanjaFastConversion": local_HangulHanjaFastConversion, "CheckHangulEndings": local_CheckHangulEndings, "EnableHangulHanjaRecentOrdering": local_EnableHangulHanjaRecentOrdering, "MultipleWordConversionsMode": local_MultipleWordConversionsMode, "DefaultBorderColor": local_DefaultBorderColor, "AllowPixelUnits": local_AllowPixelUnits, "UseCharacterUnit": local_UseCharacterUnit, "AllowCompoundNounProcessing": local_AllowCompoundNounProcessing, "AutoKeyboardSwitching": local_AutoKeyboardSwitching, "DocumentViewDirection": local_DocumentViewDirection, "ArabicNumeral": local_ArabicNumeral, "MonthNames": local_MonthNames, "CursorMovement": local_CursorMovement, "VisualSelection": local_VisualSelection, "ShowDiacritics": local_ShowDiacritics, "ShowControlCharacters": local_ShowControlCharacters, "AddControlCharacters": local_AddControlCharacters, "AddBiDirectionalMarksWhenSavingTextFile": local_AddBiDirectionalMarksWhenSavingTextFile, "StrictInitialAlefHamza": local_StrictInitialAlefHamza, "StrictFinalYaa": local_StrictFinalYaa, "HebrewMode": local_HebrewMode, "ArabicMode": local_ArabicMode, "AllowClickAndTypeMouse": local_AllowClickAndTypeMouse, "UseGermanSpellingReform": local_UseGermanSpellingReform, "InterpretHighAnsi": local_InterpretHighAnsi, "AddHebDoubleQuote": local_AddHebDoubleQuote, "UseDiffDiacColor": local_UseDiffDiacColor, "DiacriticColorVal": local_DiacriticColorVal, "OptimizeForWord97byDefault": local_OptimizeForWord97byDefault, "LocalNetworkFile": local_LocalNetworkFile, "TypeNReplace": local_TypeNReplace, "SequenceCheck": local_SequenceCheck, "BackgroundOpen": local_BackgroundOpen, "DisableFeaturesbyDefault": local_DisableFeaturesbyDefault, "PasteAdjustWordSpacing": local_PasteAdjustWordSpacing, "PasteAdjustParagraphSpacing": local_PasteAdjustParagraphSpacing, "PasteAdjustTableFormatting": local_PasteAdjustTableFormatting, "PasteSmartStyleBehavior": local_PasteSmartStyleBehavior, "PasteMergeFromPPT": local_PasteMergeFromPPT, "PasteMergeFromXL": local_PasteMergeFromXL, "CtrlClickHyperlinkToOpen": local_CtrlClickHyperlinkToOpen, "PictureWrapType": local_PictureWrapType, "DisableFeaturesIntroducedAfterbyDefault": local_DisableFeaturesIntroducedAfterbyDefault, "PasteSmartCutPaste": local_PasteSmartCutPaste, "DisplayPasteOptions": local_DisplayPasteOptions, "PromptUpdateStyle": local_PromptUpdateStyle, "DefaultEPostageApp": local_DefaultEPostageApp, "DefaultTextEncoding": local_DefaultTextEncoding, "LabelSmartTags": local_LabelSmartTags, "DisplaySmartTagButtons": local_DisplaySmartTagButtons, "WarnBeforeSavingPrintingSendingMarkup": local_WarnBeforeSavingPrintingSendingMarkup, "StoreRSIDOnSave": local_StoreRSIDOnSave, "ShowFormatError": local_ShowFormatError, "FormatScanning": local_FormatScanning, "PasteMergeLists": local_PasteMergeLists, "AutoCreateNewDrawings": local_AutoCreateNewDrawings, "SmartParaSelection": local_SmartParaSelection, "RevisionsBalloonPrintOrientation": local_RevisionsBalloonPrintOrientation, "CommentsColor": local_CommentsColor, "PrintXMLTag": local_PrintXMLTag, "PrintBackgrounds": local_PrintBackgrounds, "AllowReadingMode": local_AllowReadingMode, "ShowMarkupOpenSave": local_ShowMarkupOpenSave, "SmartCursoring": local_SmartCursoring, "MoveToTextMark": local_MoveToTextMark, "MoveFromTextMark": local_MoveFromTextMark, "BibliographyStyle": local_BibliographyStyle, "BibliographySort": local_BibliographySort, "InsertedCellColor": local_InsertedCellColor, "DeletedCellColor": local_DeletedCellColor, "MergedCellColor": local_MergedCellColor, "SplitCellColor": local_SplitCellColor, "ShowSelectionFloaties": local_ShowSelectionFloaties, "ShowMenuFloaties": local_ShowMenuFloaties, "ShowDevTools": local_ShowDevTools, "EnableLivePreview": local_EnableLivePreview, "OMathAutoBuildUp": local_OMathAutoBuildUp, "AlwaysUseClearType": local_AlwaysUseClearType, "PasteFormatWithinDocument": local_PasteFormatWithinDocument, "PasteFormatBetweenDocuments": local_PasteFormatBetweenDocuments, "PasteFormatBetweenStyledDocuments": local_PasteFormatBetweenStyledDocuments, "PasteFormatFromExternalSource": local_PasteFormatFromExternalSource, "PasteOptionKeepBulletsAndNumbers": local_PasteOptionKeepBulletsAndNumbers, "INSKeyForOvertype": local_INSKeyForOvertype, "RepeatWord": local_RepeatWord, "FrenchReform": local_FrenchReform, "ContextualSpeller": local_ContextualSpeller, "MoveToTextColor": local_MoveToTextColor, "MoveFromTextColor": local_MoveFromTextColor, "OMathCopyLF": local_OMathCopyLF, "UseNormalStyleForList": local_UseNormalStyleForList, "AllowOpenInDraftView": local_AllowOpenInDraftView, "EnableLegacyIMEMode": local_EnableLegacyIMEMode, "DoNotPromptForConvert": local_DoNotPromptForConvert, "PrecisePositioning": local_PrecisePositioning, "UpdateStyleListBehavior": local_UpdateStyleListBehavior, "StrictTaaMarboota": local_StrictTaaMarboota, "StrictRussianE": local_StrictRussianE, "SpanishMode": local_SpanishMode, "PortugalReform": local_PortugalReform, "BrazilReform": local_BrazilReform, "UpdateFieldsWithTrackedChangesAtPrint": local_UpdateFieldsWithTrackedChangesAtPrint, "DisplayAlignmentGuides": local_DisplayAlignmentGuides, "PageAlignmentGuides": local_PageAlignmentGuides, "MarginAlignmentGuides": local_MarginAlignmentGuides, "ParagraphAlignmentGuides": local_ParagraphAlignmentGuides, "EnableLiveDrag": local_EnableLiveDrag, "UseSubPixelPositioning": local_UseSubPixelPositioning, "AlertIfNotDefault": local_AlertIfNotDefault, "EnableProofingToolsAdvertisement": local_EnableProofingToolsAdvertisement, "PreferCloudSaveLocations": local_PreferCloudSaveLocations, "SkyDriveSignInOption": local_SkyDriveSignInOption, "ExpandHeadingsOnOpen": local_ExpandHeadingsOnOpen, "UseLocalUserInfo": local_UseLocalUserInfo, "CloudSignInOption": local_CloudSignInOption, "ShowPopupAddRowColToTable": local_ShowPopupAddRowColToTable, "LiveWordCount": local_LiveWordCount, "AllowCoAuthoringOnFilesWithMacros": local_AllowCoAuthoringOnFilesWithMacros, }


# Tool: 33
@mcp.tool()
async def word_get_CustomDictionaries():
	this_Global = EnsureWord()
	retVal = this_Global.CustomDictionaries
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	try:
		local_Maximum = retVal.Maximum
	except:
		local_Maximum = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Dictionaries", "Count": local_Count, "Maximum": local_Maximum, }


# Tool: 34
@mcp.tool()
async def word_get_ShowVisualBasicEditor():
	this_Global = EnsureWord()
	retVal = this_Global.ShowVisualBasicEditor
	return retVal


# Tool: 35
@mcp.tool()
async def word_get_HangulHanjaDictionaries():
	this_Global = EnsureWord()
	retVal = this_Global.HangulHanjaDictionaries
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	try:
		local_Maximum = retVal.Maximum
	except:
		local_Maximum = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "HangulHanjaConversionDictionaries", "Count": local_Count, "Maximum": local_Maximum, }


# Tool: 36
@mcp.tool()
async def word_Repeat(Times):
	"""This tool calls the Repeat method.
	
	Parameters:
		Times: the Times as VT_VARIANT
	"""
	Times = tryParseString(Times)
	this_Global = EnsureWord()
	retVal = this_Global.Repeat(Times)
	return retVal


# Tool: 37
@mcp.tool()
async def word_DDEExecute(Channel: int, Command: str):
	"""This tool calls the DDEExecute method.
	
	Parameters:
		Channel: the Channel as int
		Command: the Command as str
	"""
	this_Global = EnsureWord()
	this_Global.DDEExecute(Channel, Command)


# Tool: 38
@mcp.tool()
async def word_DDEInitiate(App: str, Topic: str):
	"""This tool calls the DDEInitiate method.
	
	Parameters:
		App: the App as str
		Topic: the Topic as str
	"""
	this_Global = EnsureWord()
	retVal = this_Global.DDEInitiate(App, Topic)
	return retVal


# Tool: 39
@mcp.tool()
async def word_DDEPoke(Channel: int, Item: str, Data: str):
	"""This tool calls the DDEPoke method.
	
	Parameters:
		Channel: the Channel as int
		Item: the Item as str
		Data: the Data as str
	"""
	this_Global = EnsureWord()
	this_Global.DDEPoke(Channel, Item, Data)


# Tool: 40
@mcp.tool()
async def word_DDERequest(Channel: int, Item: str):
	"""This tool calls the DDERequest method.
	
	Parameters:
		Channel: the Channel as int
		Item: the Item as str
	"""
	this_Global = EnsureWord()
	retVal = this_Global.DDERequest(Channel, Item)
	return retVal


# Tool: 41
@mcp.tool()
async def word_DDETerminate(Channel: int):
	"""This tool calls the DDETerminate method.
	
	Parameters:
		Channel: the Channel as int
	"""
	this_Global = EnsureWord()
	this_Global.DDETerminate(Channel)


# Tool: 42
@mcp.tool()
async def word_DDETerminateAll():
	"""This tool calls the DDETerminateAll method.
"""
	this_Global = EnsureWord()
	this_Global.DDETerminateAll()


# Tool: 43
@mcp.tool()
async def word_BuildKeyCode(Arg1: int, Arg2, Arg3, Arg4):
	"""This tool calls the BuildKeyCode method.
	
	Parameters:
		Arg1: the Arg1 as WdKey
		Arg2: the Arg2 as VT_VARIANT
		Arg3: the Arg3 as VT_VARIANT
		Arg4: the Arg4 as VT_VARIANT
	"""
	Arg2 = tryParseString(Arg2)
	Arg3 = tryParseString(Arg3)
	Arg4 = tryParseString(Arg4)
	this_Global = EnsureWord()
	retVal = this_Global.BuildKeyCode(Arg1, Arg2, Arg3, Arg4)
	return retVal


# Tool: 44
@mcp.tool()
async def word_KeyString(KeyCode: int, KeyCode2):
	"""This tool calls the KeyString method.
	
	Parameters:
		KeyCode: the KeyCode as int
		KeyCode2: the KeyCode2 as VT_VARIANT
	"""
	KeyCode2 = tryParseString(KeyCode2)
	this_Global = EnsureWord()
	retVal = this_Global.KeyString(KeyCode, KeyCode2)
	return retVal


# Tool: 45
@mcp.tool()
async def word_CheckSpelling(Word: str, CustomDictionary, IgnoreUppercase, MainDictionary, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10):
	"""This tool calls the CheckSpelling method.
	
	Parameters:
		Word: the Word as str
		CustomDictionary: the CustomDictionary as VT_VARIANT
		IgnoreUppercase: the IgnoreUppercase as VT_VARIANT
		MainDictionary: the MainDictionary as VT_VARIANT
		CustomDictionary2: the CustomDictionary2 as VT_VARIANT
		CustomDictionary3: the CustomDictionary3 as VT_VARIANT
		CustomDictionary4: the CustomDictionary4 as VT_VARIANT
		CustomDictionary5: the CustomDictionary5 as VT_VARIANT
		CustomDictionary6: the CustomDictionary6 as VT_VARIANT
		CustomDictionary7: the CustomDictionary7 as VT_VARIANT
		CustomDictionary8: the CustomDictionary8 as VT_VARIANT
		CustomDictionary9: the CustomDictionary9 as VT_VARIANT
		CustomDictionary10: the CustomDictionary10 as VT_VARIANT
	"""
	CustomDictionary = tryParseString(CustomDictionary)
	IgnoreUppercase = tryParseString(IgnoreUppercase)
	MainDictionary = tryParseString(MainDictionary)
	CustomDictionary2 = tryParseString(CustomDictionary2)
	CustomDictionary3 = tryParseString(CustomDictionary3)
	CustomDictionary4 = tryParseString(CustomDictionary4)
	CustomDictionary5 = tryParseString(CustomDictionary5)
	CustomDictionary6 = tryParseString(CustomDictionary6)
	CustomDictionary7 = tryParseString(CustomDictionary7)
	CustomDictionary8 = tryParseString(CustomDictionary8)
	CustomDictionary9 = tryParseString(CustomDictionary9)
	CustomDictionary10 = tryParseString(CustomDictionary10)
	this_Global = EnsureWord()
	retVal = this_Global.CheckSpelling(Word, CustomDictionary, IgnoreUppercase, MainDictionary, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10)
	return retVal


# Tool: 46
@mcp.tool()
async def word_GetSpellingSuggestions(Word: str, CustomDictionary, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10):
	"""This tool calls the GetSpellingSuggestions method.
	
	Parameters:
		Word: the Word as str
		CustomDictionary: the CustomDictionary as VT_VARIANT
		IgnoreUppercase: the IgnoreUppercase as VT_VARIANT
		MainDictionary: the MainDictionary as VT_VARIANT
		SuggestionMode: the SuggestionMode as VT_VARIANT
		CustomDictionary2: the CustomDictionary2 as VT_VARIANT
		CustomDictionary3: the CustomDictionary3 as VT_VARIANT
		CustomDictionary4: the CustomDictionary4 as VT_VARIANT
		CustomDictionary5: the CustomDictionary5 as VT_VARIANT
		CustomDictionary6: the CustomDictionary6 as VT_VARIANT
		CustomDictionary7: the CustomDictionary7 as VT_VARIANT
		CustomDictionary8: the CustomDictionary8 as VT_VARIANT
		CustomDictionary9: the CustomDictionary9 as VT_VARIANT
		CustomDictionary10: the CustomDictionary10 as VT_VARIANT
	"""
	CustomDictionary = tryParseString(CustomDictionary)
	IgnoreUppercase = tryParseString(IgnoreUppercase)
	MainDictionary = tryParseString(MainDictionary)
	SuggestionMode = tryParseString(SuggestionMode)
	CustomDictionary2 = tryParseString(CustomDictionary2)
	CustomDictionary3 = tryParseString(CustomDictionary3)
	CustomDictionary4 = tryParseString(CustomDictionary4)
	CustomDictionary5 = tryParseString(CustomDictionary5)
	CustomDictionary6 = tryParseString(CustomDictionary6)
	CustomDictionary7 = tryParseString(CustomDictionary7)
	CustomDictionary8 = tryParseString(CustomDictionary8)
	CustomDictionary9 = tryParseString(CustomDictionary9)
	CustomDictionary10 = tryParseString(CustomDictionary10)
	this_Global = EnsureWord()
	retVal = this_Global.GetSpellingSuggestions(Word, CustomDictionary, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10)
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	try:
		local_SpellingErrorType = retVal.SpellingErrorType
	except:
		local_SpellingErrorType = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "SpellingSuggestions", "Count": local_Count, "SpellingErrorType": local_SpellingErrorType, }


# Tool: 47
@mcp.tool()
async def word_Help(HelpType):
	"""This tool calls the Help method.
	
	Parameters:
		HelpType: the HelpType as VT_VARIANT
	"""
	HelpType = tryParseString(HelpType)
	this_Global = EnsureWord()
	this_Global.Help(HelpType)


# Tool: 48
@mcp.tool()
async def word_NewWindow():
	"""This tool calls the NewWindow method.
"""
	this_Global = EnsureWord()
	retVal = this_Global.NewWindow()
	try:
		local_Left = retVal.Left
	except:
		local_Left = None
	try:
		local_Top = retVal.Top
	except:
		local_Top = None
	try:
		local_Width = retVal.Width
	except:
		local_Width = None
	try:
		local_Height = retVal.Height
	except:
		local_Height = None
	try:
		local_Split = retVal.Split
	except:
		local_Split = None
	try:
		local_SplitVertical = retVal.SplitVertical
	except:
		local_SplitVertical = None
	try:
		local_Caption = retVal.Caption
	except:
		local_Caption = None
	try:
		local_WindowState = retVal.WindowState
	except:
		local_WindowState = None
	try:
		local_DisplayRulers = retVal.DisplayRulers
	except:
		local_DisplayRulers = None
	try:
		local_DisplayVerticalRuler = retVal.DisplayVerticalRuler
	except:
		local_DisplayVerticalRuler = None
	try:
		local_Type = retVal.Type
	except:
		local_Type = None
	try:
		local_WindowNumber = retVal.WindowNumber
	except:
		local_WindowNumber = None
	try:
		local_DisplayVerticalScrollBar = retVal.DisplayVerticalScrollBar
	except:
		local_DisplayVerticalScrollBar = None
	try:
		local_DisplayHorizontalScrollBar = retVal.DisplayHorizontalScrollBar
	except:
		local_DisplayHorizontalScrollBar = None
	try:
		local_StyleAreaWidth = retVal.StyleAreaWidth
	except:
		local_StyleAreaWidth = None
	try:
		local_DisplayScreenTips = retVal.DisplayScreenTips
	except:
		local_DisplayScreenTips = None
	try:
		local_HorizontalPercentScrolled = retVal.HorizontalPercentScrolled
	except:
		local_HorizontalPercentScrolled = None
	try:
		local_VerticalPercentScrolled = retVal.VerticalPercentScrolled
	except:
		local_VerticalPercentScrolled = None
	try:
		local_DocumentMap = retVal.DocumentMap
	except:
		local_DocumentMap = None
	try:
		local_Active = retVal.Active
	except:
		local_Active = None
	try:
		local_DocumentMapPercentWidth = retVal.DocumentMapPercentWidth
	except:
		local_DocumentMapPercentWidth = None
	try:
		local_Index = retVal.Index
	except:
		local_Index = None
	try:
		local_IMEMode = retVal.IMEMode
	except:
		local_IMEMode = None
	try:
		local_UsableWidth = retVal.UsableWidth
	except:
		local_UsableWidth = None
	try:
		local_UsableHeight = retVal.UsableHeight
	except:
		local_UsableHeight = None
	try:
		local_EnvelopeVisible = retVal.EnvelopeVisible
	except:
		local_EnvelopeVisible = None
	try:
		local_DisplayRightRuler = retVal.DisplayRightRuler
	except:
		local_DisplayRightRuler = None
	try:
		local_DisplayLeftScrollBar = retVal.DisplayLeftScrollBar
	except:
		local_DisplayLeftScrollBar = None
	try:
		local_Visible = retVal.Visible
	except:
		local_Visible = None
	try:
		local_Thumbnails = retVal.Thumbnails
	except:
		local_Thumbnails = None
	try:
		local_ShowSourceDocuments = retVal.ShowSourceDocuments
	except:
		local_ShowSourceDocuments = None
	try:
		local_Hwnd = retVal.Hwnd
	except:
		local_Hwnd = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Window", "Left": local_Left, "Top": local_Top, "Width": local_Width, "Height": local_Height, "Split": local_Split, "SplitVertical": local_SplitVertical, "Caption": local_Caption, "WindowState": local_WindowState, "DisplayRulers": local_DisplayRulers, "DisplayVerticalRuler": local_DisplayVerticalRuler, "Type": local_Type, "WindowNumber": local_WindowNumber, "DisplayVerticalScrollBar": local_DisplayVerticalScrollBar, "DisplayHorizontalScrollBar": local_DisplayHorizontalScrollBar, "StyleAreaWidth": local_StyleAreaWidth, "DisplayScreenTips": local_DisplayScreenTips, "HorizontalPercentScrolled": local_HorizontalPercentScrolled, "VerticalPercentScrolled": local_VerticalPercentScrolled, "DocumentMap": local_DocumentMap, "Active": local_Active, "DocumentMapPercentWidth": local_DocumentMapPercentWidth, "Index": local_Index, "IMEMode": local_IMEMode, "UsableWidth": local_UsableWidth, "UsableHeight": local_UsableHeight, "EnvelopeVisible": local_EnvelopeVisible, "DisplayRightRuler": local_DisplayRightRuler, "DisplayLeftScrollBar": local_DisplayLeftScrollBar, "Visible": local_Visible, "Thumbnails": local_Thumbnails, "ShowSourceDocuments": local_ShowSourceDocuments, "Hwnd": local_Hwnd, }


# Tool: 49
@mcp.tool()
async def word_CleanString(String: str):
	"""This tool calls the CleanString method.
	
	Parameters:
		String: the String as str
	"""
	this_Global = EnsureWord()
	retVal = this_Global.CleanString(String)
	return retVal


# Tool: 50
@mcp.tool()
async def word_ChangeFileOpenDirectory(Path: str):
	"""This tool calls the ChangeFileOpenDirectory method.
	
	Parameters:
		Path: the Path as str
	"""
	this_Global = EnsureWord()
	this_Global.ChangeFileOpenDirectory(Path)


# Tool: 51
@mcp.tool()
async def word_InchesToPoints(Inches: float):
	"""This tool calls the InchesToPoints method.
	
	Parameters:
		Inches: the Inches as float
	"""
	this_Global = EnsureWord()
	retVal = this_Global.InchesToPoints(Inches)
	return retVal


# Tool: 52
@mcp.tool()
async def word_CentimetersToPoints(Centimeters: float):
	"""This tool calls the CentimetersToPoints method.
	
	Parameters:
		Centimeters: the Centimeters as float
	"""
	this_Global = EnsureWord()
	retVal = this_Global.CentimetersToPoints(Centimeters)
	return retVal


# Tool: 53
@mcp.tool()
async def word_MillimetersToPoints(Millimeters: float):
	"""This tool calls the MillimetersToPoints method.
	
	Parameters:
		Millimeters: the Millimeters as float
	"""
	this_Global = EnsureWord()
	retVal = this_Global.MillimetersToPoints(Millimeters)
	return retVal


# Tool: 54
@mcp.tool()
async def word_PicasToPoints(Picas: float):
	"""This tool calls the PicasToPoints method.
	
	Parameters:
		Picas: the Picas as float
	"""
	this_Global = EnsureWord()
	retVal = this_Global.PicasToPoints(Picas)
	return retVal


# Tool: 55
@mcp.tool()
async def word_LinesToPoints(Lines: float):
	"""This tool calls the LinesToPoints method.
	
	Parameters:
		Lines: the Lines as float
	"""
	this_Global = EnsureWord()
	retVal = this_Global.LinesToPoints(Lines)
	return retVal


# Tool: 56
@mcp.tool()
async def word_PointsToInches(Points: float):
	"""This tool calls the PointsToInches method.
	
	Parameters:
		Points: the Points as float
	"""
	this_Global = EnsureWord()
	retVal = this_Global.PointsToInches(Points)
	return retVal


# Tool: 57
@mcp.tool()
async def word_PointsToCentimeters(Points: float):
	"""This tool calls the PointsToCentimeters method.
	
	Parameters:
		Points: the Points as float
	"""
	this_Global = EnsureWord()
	retVal = this_Global.PointsToCentimeters(Points)
	return retVal


# Tool: 58
@mcp.tool()
async def word_PointsToMillimeters(Points: float):
	"""This tool calls the PointsToMillimeters method.
	
	Parameters:
		Points: the Points as float
	"""
	this_Global = EnsureWord()
	retVal = this_Global.PointsToMillimeters(Points)
	return retVal


# Tool: 59
@mcp.tool()
async def word_PointsToPicas(Points: float):
	"""This tool calls the PointsToPicas method.
	
	Parameters:
		Points: the Points as float
	"""
	this_Global = EnsureWord()
	retVal = this_Global.PointsToPicas(Points)
	return retVal


# Tool: 60
@mcp.tool()
async def word_PointsToLines(Points: float):
	"""This tool calls the PointsToLines method.
	
	Parameters:
		Points: the Points as float
	"""
	this_Global = EnsureWord()
	retVal = this_Global.PointsToLines(Points)
	return retVal


# Tool: 61
@mcp.tool()
async def word_PointsToPixels(Points: float, fVertical):
	"""This tool calls the PointsToPixels method.
	
	Parameters:
		Points: the Points as float
		fVertical: the fVertical as VT_VARIANT
	"""
	fVertical = tryParseString(fVertical)
	this_Global = EnsureWord()
	retVal = this_Global.PointsToPixels(Points, fVertical)
	return retVal


# Tool: 62
@mcp.tool()
async def word_PixelsToPoints(Pixels: float, fVertical):
	"""This tool calls the PixelsToPoints method.
	
	Parameters:
		Pixels: the Pixels as float
		fVertical: the fVertical as VT_VARIANT
	"""
	fVertical = tryParseString(fVertical)
	this_Global = EnsureWord()
	retVal = this_Global.PixelsToPoints(Pixels, fVertical)
	return retVal


# Tool: 63
@mcp.tool()
async def word_get_LanguageSettings():
	this_Global = EnsureWord()
	retVal = this_Global.LanguageSettings
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "LanguageSettings"}


# Tool: 64
@mcp.tool()
async def word_get_AnswerWizard():
	this_Global = EnsureWord()
	retVal = this_Global.AnswerWizard
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "AnswerWizard"}


# Tool: 65
@mcp.tool()
async def word_get_AutoCorrectEmail():
	this_Global = EnsureWord()
	retVal = this_Global.AutoCorrectEmail
	try:
		local_CorrectDays = retVal.CorrectDays
	except:
		local_CorrectDays = None
	try:
		local_CorrectInitialCaps = retVal.CorrectInitialCaps
	except:
		local_CorrectInitialCaps = None
	try:
		local_CorrectSentenceCaps = retVal.CorrectSentenceCaps
	except:
		local_CorrectSentenceCaps = None
	try:
		local_ReplaceText = retVal.ReplaceText
	except:
		local_ReplaceText = None
	try:
		local_FirstLetterAutoAdd = retVal.FirstLetterAutoAdd
	except:
		local_FirstLetterAutoAdd = None
	try:
		local_TwoInitialCapsAutoAdd = retVal.TwoInitialCapsAutoAdd
	except:
		local_TwoInitialCapsAutoAdd = None
	try:
		local_CorrectCapsLock = retVal.CorrectCapsLock
	except:
		local_CorrectCapsLock = None
	try:
		local_CorrectHangulAndAlphabet = retVal.CorrectHangulAndAlphabet
	except:
		local_CorrectHangulAndAlphabet = None
	try:
		local_HangulAndAlphabetAutoAdd = retVal.HangulAndAlphabetAutoAdd
	except:
		local_HangulAndAlphabetAutoAdd = None
	try:
		local_ReplaceTextFromSpellingChecker = retVal.ReplaceTextFromSpellingChecker
	except:
		local_ReplaceTextFromSpellingChecker = None
	try:
		local_OtherCorrectionsAutoAdd = retVal.OtherCorrectionsAutoAdd
	except:
		local_OtherCorrectionsAutoAdd = None
	try:
		local_CorrectKeyboardSetting = retVal.CorrectKeyboardSetting
	except:
		local_CorrectKeyboardSetting = None
	try:
		local_CorrectTableCells = retVal.CorrectTableCells
	except:
		local_CorrectTableCells = None
	try:
		local_DisplayAutoCorrectOptions = retVal.DisplayAutoCorrectOptions
	except:
		local_DisplayAutoCorrectOptions = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "AutoCorrect", "CorrectDays": local_CorrectDays, "CorrectInitialCaps": local_CorrectInitialCaps, "CorrectSentenceCaps": local_CorrectSentenceCaps, "ReplaceText": local_ReplaceText, "FirstLetterAutoAdd": local_FirstLetterAutoAdd, "TwoInitialCapsAutoAdd": local_TwoInitialCapsAutoAdd, "CorrectCapsLock": local_CorrectCapsLock, "CorrectHangulAndAlphabet": local_CorrectHangulAndAlphabet, "HangulAndAlphabetAutoAdd": local_HangulAndAlphabetAutoAdd, "ReplaceTextFromSpellingChecker": local_ReplaceTextFromSpellingChecker, "OtherCorrectionsAutoAdd": local_OtherCorrectionsAutoAdd, "CorrectKeyboardSetting": local_CorrectKeyboardSetting, "CorrectTableCells": local_CorrectTableCells, "DisplayAutoCorrectOptions": local_DisplayAutoCorrectOptions, }


# Tool: 66
@mcp.tool()
async def word_get_ProtectedViewWindows():
	this_Global = EnsureWord()
	retVal = this_Global.ProtectedViewWindows
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ProtectedViewWindows", "Count": local_Count, }


# Tool: 67
@mcp.tool()
async def word_get_ActiveProtectedViewWindow():
	this_Global = EnsureWord()
	retVal = this_Global.ActiveProtectedViewWindow
	try:
		local_Caption = retVal.Caption
	except:
		local_Caption = None
	try:
		local_Left = retVal.Left
	except:
		local_Left = None
	try:
		local_Top = retVal.Top
	except:
		local_Top = None
	try:
		local_Width = retVal.Width
	except:
		local_Width = None
	try:
		local_Height = retVal.Height
	except:
		local_Height = None
	try:
		local_WindowState = retVal.WindowState
	except:
		local_WindowState = None
	try:
		local_Active = retVal.Active
	except:
		local_Active = None
	try:
		local_Index = retVal.Index
	except:
		local_Index = None
	try:
		local_Visible = retVal.Visible
	except:
		local_Visible = None
	try:
		local_SourceName = retVal.SourceName
	except:
		local_SourceName = None
	try:
		local_SourcePath = retVal.SourcePath
	except:
		local_SourcePath = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ProtectedViewWindow", "Caption": local_Caption, "Left": local_Left, "Top": local_Top, "Width": local_Width, "Height": local_Height, "WindowState": local_WindowState, "Active": local_Active, "Index": local_Index, "Visible": local_Visible, "SourceName": local_SourceName, "SourcePath": local_SourcePath, }


# Tool: 68
@mcp.tool()
async def word_get_IsSandboxed():
	this_Global = EnsureWord()
	retVal = this_Global.IsSandboxed
	return retVal


# Tool: 69
@mcp.tool()
async def word_Documents_Item(this_Documents_wordObjId: str, Index):
	"""This tool calls the Item methodon an Documents object. Pass the __WordObjectId of Documents of the object you want to call the method on as the first parameter
	
	Parameters:
		Index: the Index as VT_VARIANT
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	Index = tryParseString(Index)
	retVal = this_Documents.Item(Index)
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Document"}


# Tool: 70
@mcp.tool()
async def word_Documents_Close(this_Documents_wordObjId: str, SaveChanges, OriginalFormat, RouteDocument):
	"""This tool calls the Close methodon an Documents object. Pass the __WordObjectId of Documents of the object you want to call the method on as the first parameter
	
	Parameters:
		SaveChanges: the SaveChanges as VT_VARIANT
		OriginalFormat: the OriginalFormat as VT_VARIANT
		RouteDocument: the RouteDocument as VT_VARIANT
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	SaveChanges = tryParseString(SaveChanges)
	OriginalFormat = tryParseString(OriginalFormat)
	RouteDocument = tryParseString(RouteDocument)
	this_Documents.Close(SaveChanges, OriginalFormat, RouteDocument)


# Tool: 71
@mcp.tool()
async def word_Documents_AddOld(this_Documents_wordObjId: str, Template, NewTemplate):
	"""This tool calls the AddOld methodon an Documents object. Pass the __WordObjectId of Documents of the object you want to call the method on as the first parameter
	
	Parameters:
		Template: the Template as VT_VARIANT
		NewTemplate: the NewTemplate as VT_VARIANT
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	Template = tryParseString(Template)
	NewTemplate = tryParseString(NewTemplate)
	retVal = this_Documents.AddOld(Template, NewTemplate)
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Document"}


# Tool: 72
@mcp.tool()
async def word_Documents_OpenOld(this_Documents_wordObjId: str, FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format):
	"""This tool calls the OpenOld methodon an Documents object. Pass the __WordObjectId of Documents of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as VT_VARIANT
		ConfirmConversions: the ConfirmConversions as VT_VARIANT
		ReadOnly: the ReadOnly as VT_VARIANT
		AddToRecentFiles: the AddToRecentFiles as VT_VARIANT
		PasswordDocument: the PasswordDocument as VT_VARIANT
		PasswordTemplate: the PasswordTemplate as VT_VARIANT
		Revert: the Revert as VT_VARIANT
		WritePasswordDocument: the WritePasswordDocument as VT_VARIANT
		WritePasswordTemplate: the WritePasswordTemplate as VT_VARIANT
		Format: the Format as VT_VARIANT
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	FileName = tryParseString(FileName)
	ConfirmConversions = tryParseString(ConfirmConversions)
	ReadOnly = tryParseString(ReadOnly)
	AddToRecentFiles = tryParseString(AddToRecentFiles)
	PasswordDocument = tryParseString(PasswordDocument)
	PasswordTemplate = tryParseString(PasswordTemplate)
	Revert = tryParseString(Revert)
	WritePasswordDocument = tryParseString(WritePasswordDocument)
	WritePasswordTemplate = tryParseString(WritePasswordTemplate)
	Format = tryParseString(Format)
	retVal = this_Documents.OpenOld(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format)
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Document"}


# Tool: 73
@mcp.tool()
async def word_Documents_Save(this_Documents_wordObjId: str, NoPrompt, OriginalFormat):
	"""This tool calls the Save methodon an Documents object. Pass the __WordObjectId of Documents of the object you want to call the method on as the first parameter
	
	Parameters:
		NoPrompt: the NoPrompt as VT_VARIANT
		OriginalFormat: the OriginalFormat as VT_VARIANT
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	NoPrompt = tryParseString(NoPrompt)
	OriginalFormat = tryParseString(OriginalFormat)
	this_Documents.Save(NoPrompt, OriginalFormat)


# Tool: 74
@mcp.tool()
async def word_Documents_Add(this_Documents_wordObjId: str, Template, NewTemplate, DocumentType, Visible):
	"""This tool calls the Add methodon an Documents object. Pass the __WordObjectId of Documents of the object you want to call the method on as the first parameter
	
	Parameters:
		Template: the Template as VT_VARIANT
		NewTemplate: the NewTemplate as VT_VARIANT
		DocumentType: the DocumentType as VT_VARIANT
		Visible: the Visible as VT_VARIANT
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	Template = tryParseString(Template)
	NewTemplate = tryParseString(NewTemplate)
	DocumentType = tryParseString(DocumentType)
	Visible = tryParseString(Visible)
	retVal = this_Documents.Add(Template, NewTemplate, DocumentType, Visible)
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Document"}


# Tool: 75
@mcp.tool()
async def word_Documents_Open2000(this_Documents_wordObjId: str, FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible):
	"""This tool calls the Open2000 methodon an Documents object. Pass the __WordObjectId of Documents of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as VT_VARIANT
		ConfirmConversions: the ConfirmConversions as VT_VARIANT
		ReadOnly: the ReadOnly as VT_VARIANT
		AddToRecentFiles: the AddToRecentFiles as VT_VARIANT
		PasswordDocument: the PasswordDocument as VT_VARIANT
		PasswordTemplate: the PasswordTemplate as VT_VARIANT
		Revert: the Revert as VT_VARIANT
		WritePasswordDocument: the WritePasswordDocument as VT_VARIANT
		WritePasswordTemplate: the WritePasswordTemplate as VT_VARIANT
		Format: the Format as VT_VARIANT
		Encoding: the Encoding as VT_VARIANT
		Visible: the Visible as VT_VARIANT
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	FileName = tryParseString(FileName)
	ConfirmConversions = tryParseString(ConfirmConversions)
	ReadOnly = tryParseString(ReadOnly)
	AddToRecentFiles = tryParseString(AddToRecentFiles)
	PasswordDocument = tryParseString(PasswordDocument)
	PasswordTemplate = tryParseString(PasswordTemplate)
	Revert = tryParseString(Revert)
	WritePasswordDocument = tryParseString(WritePasswordDocument)
	WritePasswordTemplate = tryParseString(WritePasswordTemplate)
	Format = tryParseString(Format)
	Encoding = tryParseString(Encoding)
	Visible = tryParseString(Visible)
	retVal = this_Documents.Open2000(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible)
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Document"}


# Tool: 76
@mcp.tool()
async def word_Documents_CheckOut(this_Documents_wordObjId: str, FileName: str):
	"""This tool calls the CheckOut methodon an Documents object. Pass the __WordObjectId of Documents of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as str
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	this_Documents.CheckOut(FileName)


# Tool: 77
@mcp.tool()
async def word_Documents_CanCheckOut(this_Documents_wordObjId: str, FileName: str):
	"""This tool calls the CanCheckOut methodon an Documents object. Pass the __WordObjectId of Documents of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as str
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	retVal = this_Documents.CanCheckOut(FileName)
	return retVal


# Tool: 78
@mcp.tool()
async def word_Documents_Open2002(this_Documents_wordObjId: str, FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog):
	"""This tool calls the Open2002 methodon an Documents object. Pass the __WordObjectId of Documents of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as VT_VARIANT
		ConfirmConversions: the ConfirmConversions as VT_VARIANT
		ReadOnly: the ReadOnly as VT_VARIANT
		AddToRecentFiles: the AddToRecentFiles as VT_VARIANT
		PasswordDocument: the PasswordDocument as VT_VARIANT
		PasswordTemplate: the PasswordTemplate as VT_VARIANT
		Revert: the Revert as VT_VARIANT
		WritePasswordDocument: the WritePasswordDocument as VT_VARIANT
		WritePasswordTemplate: the WritePasswordTemplate as VT_VARIANT
		Format: the Format as VT_VARIANT
		Encoding: the Encoding as VT_VARIANT
		Visible: the Visible as VT_VARIANT
		OpenAndRepair: the OpenAndRepair as VT_VARIANT
		DocumentDirection: the DocumentDirection as VT_VARIANT
		NoEncodingDialog: the NoEncodingDialog as VT_VARIANT
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	FileName = tryParseString(FileName)
	ConfirmConversions = tryParseString(ConfirmConversions)
	ReadOnly = tryParseString(ReadOnly)
	AddToRecentFiles = tryParseString(AddToRecentFiles)
	PasswordDocument = tryParseString(PasswordDocument)
	PasswordTemplate = tryParseString(PasswordTemplate)
	Revert = tryParseString(Revert)
	WritePasswordDocument = tryParseString(WritePasswordDocument)
	WritePasswordTemplate = tryParseString(WritePasswordTemplate)
	Format = tryParseString(Format)
	Encoding = tryParseString(Encoding)
	Visible = tryParseString(Visible)
	OpenAndRepair = tryParseString(OpenAndRepair)
	DocumentDirection = tryParseString(DocumentDirection)
	NoEncodingDialog = tryParseString(NoEncodingDialog)
	retVal = this_Documents.Open2002(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog)
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Document"}


# Tool: 79
@mcp.tool()
async def word_Documents_Open(this_Documents_wordObjId: str, FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog, XMLTransform):
	"""This tool calls the Open methodon an Documents object. Pass the __WordObjectId of Documents of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as VT_VARIANT
		ConfirmConversions: the ConfirmConversions as VT_VARIANT
		ReadOnly: the ReadOnly as VT_VARIANT
		AddToRecentFiles: the AddToRecentFiles as VT_VARIANT
		PasswordDocument: the PasswordDocument as VT_VARIANT
		PasswordTemplate: the PasswordTemplate as VT_VARIANT
		Revert: the Revert as VT_VARIANT
		WritePasswordDocument: the WritePasswordDocument as VT_VARIANT
		WritePasswordTemplate: the WritePasswordTemplate as VT_VARIANT
		Format: the Format as VT_VARIANT
		Encoding: the Encoding as VT_VARIANT
		Visible: the Visible as VT_VARIANT
		OpenAndRepair: the OpenAndRepair as VT_VARIANT
		DocumentDirection: the DocumentDirection as VT_VARIANT
		NoEncodingDialog: the NoEncodingDialog as VT_VARIANT
		XMLTransform: the XMLTransform as VT_VARIANT
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	FileName = tryParseString(FileName)
	ConfirmConversions = tryParseString(ConfirmConversions)
	ReadOnly = tryParseString(ReadOnly)
	AddToRecentFiles = tryParseString(AddToRecentFiles)
	PasswordDocument = tryParseString(PasswordDocument)
	PasswordTemplate = tryParseString(PasswordTemplate)
	Revert = tryParseString(Revert)
	WritePasswordDocument = tryParseString(WritePasswordDocument)
	WritePasswordTemplate = tryParseString(WritePasswordTemplate)
	Format = tryParseString(Format)
	Encoding = tryParseString(Encoding)
	Visible = tryParseString(Visible)
	OpenAndRepair = tryParseString(OpenAndRepair)
	DocumentDirection = tryParseString(DocumentDirection)
	NoEncodingDialog = tryParseString(NoEncodingDialog)
	XMLTransform = tryParseString(XMLTransform)
	retVal = this_Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog, XMLTransform)
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Document"}


# Tool: 80
@mcp.tool()
async def word_Documents_OpenNoRepairDialog(this_Documents_wordObjId: str, FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog, XMLTransform):
	"""This tool calls the OpenNoRepairDialog methodon an Documents object. Pass the __WordObjectId of Documents of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as VT_VARIANT
		ConfirmConversions: the ConfirmConversions as VT_VARIANT
		ReadOnly: the ReadOnly as VT_VARIANT
		AddToRecentFiles: the AddToRecentFiles as VT_VARIANT
		PasswordDocument: the PasswordDocument as VT_VARIANT
		PasswordTemplate: the PasswordTemplate as VT_VARIANT
		Revert: the Revert as VT_VARIANT
		WritePasswordDocument: the WritePasswordDocument as VT_VARIANT
		WritePasswordTemplate: the WritePasswordTemplate as VT_VARIANT
		Format: the Format as VT_VARIANT
		Encoding: the Encoding as VT_VARIANT
		Visible: the Visible as VT_VARIANT
		OpenAndRepair: the OpenAndRepair as VT_VARIANT
		DocumentDirection: the DocumentDirection as VT_VARIANT
		NoEncodingDialog: the NoEncodingDialog as VT_VARIANT
		XMLTransform: the XMLTransform as VT_VARIANT
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	FileName = tryParseString(FileName)
	ConfirmConversions = tryParseString(ConfirmConversions)
	ReadOnly = tryParseString(ReadOnly)
	AddToRecentFiles = tryParseString(AddToRecentFiles)
	PasswordDocument = tryParseString(PasswordDocument)
	PasswordTemplate = tryParseString(PasswordTemplate)
	Revert = tryParseString(Revert)
	WritePasswordDocument = tryParseString(WritePasswordDocument)
	WritePasswordTemplate = tryParseString(WritePasswordTemplate)
	Format = tryParseString(Format)
	Encoding = tryParseString(Encoding)
	Visible = tryParseString(Visible)
	OpenAndRepair = tryParseString(OpenAndRepair)
	DocumentDirection = tryParseString(DocumentDirection)
	NoEncodingDialog = tryParseString(NoEncodingDialog)
	XMLTransform = tryParseString(XMLTransform)
	retVal = this_Documents.OpenNoRepairDialog(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog, XMLTransform)
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Document"}


# Tool: 81
@mcp.tool()
async def word_Documents_AddBlogDocument(this_Documents_wordObjId: str, ProviderID: str, PostURL: str, BlogName: str, PostID: str):
	"""This tool calls the AddBlogDocument methodon an Documents object. Pass the __WordObjectId of Documents of the object you want to call the method on as the first parameter
	
	Parameters:
		ProviderID: the ProviderID as str
		PostURL: the PostURL as str
		BlogName: the BlogName as str
		PostID: the PostID as str
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	retVal = this_Documents.AddBlogDocument(ProviderID, PostURL, BlogName, PostID)
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Document"}


# Tool: 82
@mcp.tool()
async def word_Documents_get_Property(this_Documents_wordObjId: str, propertyName: str):
	"""Gets properties of Documents
	
	propertyName: Name of the property. Can be one of ...
		Count
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	
	EnsureWord()
	if (propertyName == "Count"):
		retVal = this_Documents.Count
		return retVal


# Tool: 83
@mcp.tool()
async def word_Documents_set_Property(this_Documents_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Documents
	
	propertyName: Name of the property. Can be one of ...
		
	"""
	this_Documents = get_object(this_Documents_wordObjId)
	
	EnsureWord()


# Tool: 84
@mcp.tool()
async def word_Document_Close(this_Document_wordObjId: str, SaveChanges, OriginalFormat, RouteDocument):
	"""This tool calls the Close methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		SaveChanges: the SaveChanges as VT_VARIANT
		OriginalFormat: the OriginalFormat as VT_VARIANT
		RouteDocument: the RouteDocument as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	SaveChanges = tryParseString(SaveChanges)
	OriginalFormat = tryParseString(OriginalFormat)
	RouteDocument = tryParseString(RouteDocument)
	this_Document.Close(SaveChanges, OriginalFormat, RouteDocument)


# Tool: 85
@mcp.tool()
async def word_Document_SaveAs2000(this_Document_wordObjId: str, FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter):
	"""This tool calls the SaveAs2000 methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as VT_VARIANT
		FileFormat: the FileFormat as VT_VARIANT
		LockComments: the LockComments as VT_VARIANT
		Password: the Password as VT_VARIANT
		AddToRecentFiles: the AddToRecentFiles as VT_VARIANT
		WritePassword: the WritePassword as VT_VARIANT
		ReadOnlyRecommended: the ReadOnlyRecommended as VT_VARIANT
		EmbedTrueTypeFonts: the EmbedTrueTypeFonts as VT_VARIANT
		SaveNativePictureFormat: the SaveNativePictureFormat as VT_VARIANT
		SaveFormsData: the SaveFormsData as VT_VARIANT
		SaveAsAOCELetter: the SaveAsAOCELetter as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	FileName = tryParseString(FileName)
	FileFormat = tryParseString(FileFormat)
	LockComments = tryParseString(LockComments)
	Password = tryParseString(Password)
	AddToRecentFiles = tryParseString(AddToRecentFiles)
	WritePassword = tryParseString(WritePassword)
	ReadOnlyRecommended = tryParseString(ReadOnlyRecommended)
	EmbedTrueTypeFonts = tryParseString(EmbedTrueTypeFonts)
	SaveNativePictureFormat = tryParseString(SaveNativePictureFormat)
	SaveFormsData = tryParseString(SaveFormsData)
	SaveAsAOCELetter = tryParseString(SaveAsAOCELetter)
	this_Document.SaveAs2000(FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter)


# Tool: 86
@mcp.tool()
async def word_Document_Repaginate(this_Document_wordObjId: str):
	"""This tool calls the Repaginate methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.Repaginate()


# Tool: 87
@mcp.tool()
async def word_Document_FitToPages(this_Document_wordObjId: str):
	"""This tool calls the FitToPages methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.FitToPages()


# Tool: 88
@mcp.tool()
async def word_Document_ManualHyphenation(this_Document_wordObjId: str):
	"""This tool calls the ManualHyphenation methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ManualHyphenation()


# Tool: 89
@mcp.tool()
async def word_Document_Select(this_Document_wordObjId: str):
	"""This tool calls the Select methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.Select()


# Tool: 90
@mcp.tool()
async def word_Document_DataForm(this_Document_wordObjId: str):
	"""This tool calls the DataForm methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.DataForm()


# Tool: 91
@mcp.tool()
async def word_Document_Route(this_Document_wordObjId: str):
	"""This tool calls the Route methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.Route()


# Tool: 92
@mcp.tool()
async def word_Document_Save(this_Document_wordObjId: str):
	"""This tool calls the Save methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.Save()


# Tool: 93
@mcp.tool()
async def word_Document_PrintOutOld(this_Document_wordObjId: str, Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, ActivePrinterMacGX, ManualDuplexPrint):
	"""This tool calls the PrintOutOld methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Background: the Background as VT_VARIANT
		Append: the Append as VT_VARIANT
		Range: the Range as VT_VARIANT
		OutputFileName: the OutputFileName as VT_VARIANT
		From: the From as VT_VARIANT
		To: the To as VT_VARIANT
		Item: the Item as VT_VARIANT
		Copies: the Copies as VT_VARIANT
		Pages: the Pages as VT_VARIANT
		PageType: the PageType as VT_VARIANT
		PrintToFile: the PrintToFile as VT_VARIANT
		Collate: the Collate as VT_VARIANT
		ActivePrinterMacGX: the ActivePrinterMacGX as VT_VARIANT
		ManualDuplexPrint: the ManualDuplexPrint as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Background = tryParseString(Background)
	Append = tryParseString(Append)
	Range = tryParseString(Range)
	OutputFileName = tryParseString(OutputFileName)
	From = tryParseString(From)
	To = tryParseString(To)
	Item = tryParseString(Item)
	Copies = tryParseString(Copies)
	Pages = tryParseString(Pages)
	PageType = tryParseString(PageType)
	PrintToFile = tryParseString(PrintToFile)
	Collate = tryParseString(Collate)
	ActivePrinterMacGX = tryParseString(ActivePrinterMacGX)
	ManualDuplexPrint = tryParseString(ManualDuplexPrint)
	this_Document.PrintOutOld(Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, ActivePrinterMacGX, ManualDuplexPrint)


# Tool: 94
@mcp.tool()
async def word_Document_SendMail(this_Document_wordObjId: str):
	"""This tool calls the SendMail methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.SendMail()


# Tool: 95
@mcp.tool()
async def word_Document_Range(this_Document_wordObjId: str, Start, End):
	"""This tool calls the Range methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Start: the Start as VT_VARIANT
		End: the End as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Start = tryParseString(Start)
	End = tryParseString(End)
	retVal = this_Document.Range(Start, End)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 96
@mcp.tool()
async def word_Document_RunAutoMacro(this_Document_wordObjId: str, Which: int):
	"""This tool calls the RunAutoMacro methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Which: the Which as WdAutoMacros
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.RunAutoMacro(Which)


# Tool: 97
@mcp.tool()
async def word_Document_Activate(this_Document_wordObjId: str):
	"""This tool calls the Activate methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.Activate()


# Tool: 98
@mcp.tool()
async def word_Document_PrintPreview(this_Document_wordObjId: str):
	"""This tool calls the PrintPreview methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.PrintPreview()


# Tool: 99
@mcp.tool()
async def word_Document_GoTo(this_Document_wordObjId: str, What, Which, Count, Name):
	"""This tool calls the GoTo methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		What: the What as VT_VARIANT
		Which: the Which as VT_VARIANT
		Count: the Count as VT_VARIANT
		Name: the Name as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	What = tryParseString(What)
	Which = tryParseString(Which)
	Count = tryParseString(Count)
	Name = tryParseString(Name)
	retVal = this_Document.GoTo(What, Which, Count, Name)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 100
@mcp.tool()
async def word_Document_Undo(this_Document_wordObjId: str, Times):
	"""This tool calls the Undo methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Times: the Times as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Times = tryParseString(Times)
	retVal = this_Document.Undo(Times)
	return retVal


# Tool: 101
@mcp.tool()
async def word_Document_Redo(this_Document_wordObjId: str, Times):
	"""This tool calls the Redo methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Times: the Times as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Times = tryParseString(Times)
	retVal = this_Document.Redo(Times)
	return retVal


# Tool: 102
@mcp.tool()
async def word_Document_ComputeStatistics(this_Document_wordObjId: str, Statistic: int, IncludeFootnotesAndEndnotes):
	"""This tool calls the ComputeStatistics methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Statistic: the Statistic as WdStatistic
		IncludeFootnotesAndEndnotes: the IncludeFootnotesAndEndnotes as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	IncludeFootnotesAndEndnotes = tryParseString(IncludeFootnotesAndEndnotes)
	retVal = this_Document.ComputeStatistics(Statistic, IncludeFootnotesAndEndnotes)
	return retVal


# Tool: 103
@mcp.tool()
async def word_Document_MakeCompatibilityDefault(this_Document_wordObjId: str):
	"""This tool calls the MakeCompatibilityDefault methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.MakeCompatibilityDefault()


# Tool: 104
@mcp.tool()
async def word_Document_Protect2002(this_Document_wordObjId: str, Type: int, NoReset, Password):
	"""This tool calls the Protect2002 methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Type: the Type as WdProtectionType
		NoReset: the NoReset as VT_VARIANT
		Password: the Password as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	NoReset = tryParseString(NoReset)
	Password = tryParseString(Password)
	this_Document.Protect2002(Type, NoReset, Password)


# Tool: 105
@mcp.tool()
async def word_Document_Unprotect(this_Document_wordObjId: str, Password):
	"""This tool calls the Unprotect methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Password: the Password as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Password = tryParseString(Password)
	this_Document.Unprotect(Password)


# Tool: 106
@mcp.tool()
async def word_Document_EditionOptions(this_Document_wordObjId: str, Type: int, Option: int, Name: str, Format):
	"""This tool calls the EditionOptions methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Type: the Type as WdEditionType
		Option: the Option as WdEditionOption
		Name: the Name as str
		Format: the Format as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Format = tryParseString(Format)
	this_Document.EditionOptions(Type, Option, Name, Format)


# Tool: 107
@mcp.tool()
async def word_Document_RunLetterWizard(this_Document_wordObjId: str, LetterContent, WizardMode):
	"""This tool calls the RunLetterWizard methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		LetterContent: the LetterContent as VT_VARIANT
		WizardMode: the WizardMode as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	LetterContent = tryParseString(LetterContent)
	WizardMode = tryParseString(WizardMode)
	this_Document.RunLetterWizard(LetterContent, WizardMode)


# Tool: 108
@mcp.tool()
async def word_Document_GetLetterContent(this_Document_wordObjId: str):
	"""This tool calls the GetLetterContent methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	retVal = this_Document.GetLetterContent()
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "LetterContent"}


# Tool: 109
@mcp.tool()
async def word_Document_SetLetterContent(this_Document_wordObjId: str, LetterContent):
	"""This tool calls the SetLetterContent methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		LetterContent: the LetterContent as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	LetterContent = tryParseString(LetterContent)
	this_Document.SetLetterContent(LetterContent)


# Tool: 110
@mcp.tool()
async def word_Document_CopyStylesFromTemplate(this_Document_wordObjId: str, Template: str):
	"""This tool calls the CopyStylesFromTemplate methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Template: the Template as str
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.CopyStylesFromTemplate(Template)


# Tool: 111
@mcp.tool()
async def word_Document_UpdateStyles(this_Document_wordObjId: str):
	"""This tool calls the UpdateStyles methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.UpdateStyles()


# Tool: 112
@mcp.tool()
async def word_Document_CheckGrammar(this_Document_wordObjId: str):
	"""This tool calls the CheckGrammar methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.CheckGrammar()


# Tool: 113
@mcp.tool()
async def word_Document_CheckSpelling(this_Document_wordObjId: str, CustomDictionary, IgnoreUppercase, AlwaysSuggest, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10):
	"""This tool calls the CheckSpelling methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		CustomDictionary: the CustomDictionary as VT_VARIANT
		IgnoreUppercase: the IgnoreUppercase as VT_VARIANT
		AlwaysSuggest: the AlwaysSuggest as VT_VARIANT
		CustomDictionary2: the CustomDictionary2 as VT_VARIANT
		CustomDictionary3: the CustomDictionary3 as VT_VARIANT
		CustomDictionary4: the CustomDictionary4 as VT_VARIANT
		CustomDictionary5: the CustomDictionary5 as VT_VARIANT
		CustomDictionary6: the CustomDictionary6 as VT_VARIANT
		CustomDictionary7: the CustomDictionary7 as VT_VARIANT
		CustomDictionary8: the CustomDictionary8 as VT_VARIANT
		CustomDictionary9: the CustomDictionary9 as VT_VARIANT
		CustomDictionary10: the CustomDictionary10 as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	CustomDictionary = tryParseString(CustomDictionary)
	IgnoreUppercase = tryParseString(IgnoreUppercase)
	AlwaysSuggest = tryParseString(AlwaysSuggest)
	CustomDictionary2 = tryParseString(CustomDictionary2)
	CustomDictionary3 = tryParseString(CustomDictionary3)
	CustomDictionary4 = tryParseString(CustomDictionary4)
	CustomDictionary5 = tryParseString(CustomDictionary5)
	CustomDictionary6 = tryParseString(CustomDictionary6)
	CustomDictionary7 = tryParseString(CustomDictionary7)
	CustomDictionary8 = tryParseString(CustomDictionary8)
	CustomDictionary9 = tryParseString(CustomDictionary9)
	CustomDictionary10 = tryParseString(CustomDictionary10)
	this_Document.CheckSpelling(CustomDictionary, IgnoreUppercase, AlwaysSuggest, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10)


# Tool: 114
@mcp.tool()
async def word_Document_FollowHyperlink(this_Document_wordObjId: str, Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo):
	"""This tool calls the FollowHyperlink methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Address: the Address as VT_VARIANT
		SubAddress: the SubAddress as VT_VARIANT
		NewWindow: the NewWindow as VT_VARIANT
		AddHistory: the AddHistory as VT_VARIANT
		ExtraInfo: the ExtraInfo as VT_VARIANT
		Method: the Method as VT_VARIANT
		HeaderInfo: the HeaderInfo as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Address = tryParseString(Address)
	SubAddress = tryParseString(SubAddress)
	NewWindow = tryParseString(NewWindow)
	AddHistory = tryParseString(AddHistory)
	ExtraInfo = tryParseString(ExtraInfo)
	Method = tryParseString(Method)
	HeaderInfo = tryParseString(HeaderInfo)
	this_Document.FollowHyperlink(Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo)


# Tool: 115
@mcp.tool()
async def word_Document_AddToFavorites(this_Document_wordObjId: str):
	"""This tool calls the AddToFavorites methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.AddToFavorites()


# Tool: 116
@mcp.tool()
async def word_Document_Reload(this_Document_wordObjId: str):
	"""This tool calls the Reload methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.Reload()


# Tool: 117
@mcp.tool()
async def word_Document_AutoSummarize(this_Document_wordObjId: str, Length, Mode, UpdateProperties):
	"""This tool calls the AutoSummarize methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Length: the Length as VT_VARIANT
		Mode: the Mode as VT_VARIANT
		UpdateProperties: the UpdateProperties as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Length = tryParseString(Length)
	Mode = tryParseString(Mode)
	UpdateProperties = tryParseString(UpdateProperties)
	retVal = this_Document.AutoSummarize(Length, Mode, UpdateProperties)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 118
@mcp.tool()
async def word_Document_RemoveNumbers(this_Document_wordObjId: str, NumberType):
	"""This tool calls the RemoveNumbers methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		NumberType: the NumberType as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	NumberType = tryParseString(NumberType)
	this_Document.RemoveNumbers(NumberType)


# Tool: 119
@mcp.tool()
async def word_Document_ConvertNumbersToText(this_Document_wordObjId: str, NumberType):
	"""This tool calls the ConvertNumbersToText methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		NumberType: the NumberType as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	NumberType = tryParseString(NumberType)
	this_Document.ConvertNumbersToText(NumberType)


# Tool: 120
@mcp.tool()
async def word_Document_CountNumberedItems(this_Document_wordObjId: str, NumberType, Level):
	"""This tool calls the CountNumberedItems methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		NumberType: the NumberType as VT_VARIANT
		Level: the Level as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	NumberType = tryParseString(NumberType)
	Level = tryParseString(Level)
	retVal = this_Document.CountNumberedItems(NumberType, Level)
	return retVal


# Tool: 121
@mcp.tool()
async def word_Document_Post(this_Document_wordObjId: str):
	"""This tool calls the Post methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.Post()


# Tool: 122
@mcp.tool()
async def word_Document_ToggleFormsDesign(this_Document_wordObjId: str):
	"""This tool calls the ToggleFormsDesign methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ToggleFormsDesign()


# Tool: 123
@mcp.tool()
async def word_Document_Compare2000(this_Document_wordObjId: str, Name: str):
	"""This tool calls the Compare2000 methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Name: the Name as str
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.Compare2000(Name)


# Tool: 124
@mcp.tool()
async def word_Document_UpdateSummaryProperties(this_Document_wordObjId: str):
	"""This tool calls the UpdateSummaryProperties methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.UpdateSummaryProperties()


# Tool: 125
@mcp.tool()
async def word_Document_GetCrossReferenceItems(this_Document_wordObjId: str, ReferenceType):
	"""This tool calls the GetCrossReferenceItems methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		ReferenceType: the ReferenceType as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	ReferenceType = tryParseString(ReferenceType)
	retVal = this_Document.GetCrossReferenceItems(ReferenceType)
	return retVal


# Tool: 126
@mcp.tool()
async def word_Document_AutoFormat(this_Document_wordObjId: str):
	"""This tool calls the AutoFormat methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.AutoFormat()


# Tool: 127
@mcp.tool()
async def word_Document_ViewCode(this_Document_wordObjId: str):
	"""This tool calls the ViewCode methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ViewCode()


# Tool: 128
@mcp.tool()
async def word_Document_ViewPropertyBrowser(this_Document_wordObjId: str):
	"""This tool calls the ViewPropertyBrowser methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ViewPropertyBrowser()


# Tool: 129
@mcp.tool()
async def word_Document_ForwardMailer(this_Document_wordObjId: str):
	"""This tool calls the ForwardMailer methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ForwardMailer()


# Tool: 130
@mcp.tool()
async def word_Document_Reply(this_Document_wordObjId: str):
	"""This tool calls the Reply methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.Reply()


# Tool: 131
@mcp.tool()
async def word_Document_ReplyAll(this_Document_wordObjId: str):
	"""This tool calls the ReplyAll methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ReplyAll()


# Tool: 132
@mcp.tool()
async def word_Document_SendMailer(this_Document_wordObjId: str, FileFormat, Priority):
	"""This tool calls the SendMailer methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		FileFormat: the FileFormat as VT_VARIANT
		Priority: the Priority as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	FileFormat = tryParseString(FileFormat)
	Priority = tryParseString(Priority)
	this_Document.SendMailer(FileFormat, Priority)


# Tool: 133
@mcp.tool()
async def word_Document_UndoClear(this_Document_wordObjId: str):
	"""This tool calls the UndoClear methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.UndoClear()


# Tool: 134
@mcp.tool()
async def word_Document_PresentIt(this_Document_wordObjId: str):
	"""This tool calls the PresentIt methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.PresentIt()


# Tool: 135
@mcp.tool()
async def word_Document_SendFax(this_Document_wordObjId: str, Address: str, Subject):
	"""This tool calls the SendFax methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Address: the Address as str
		Subject: the Subject as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Subject = tryParseString(Subject)
	this_Document.SendFax(Address, Subject)


# Tool: 136
@mcp.tool()
async def word_Document_Merge2000(this_Document_wordObjId: str, FileName: str):
	"""This tool calls the Merge2000 methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as str
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.Merge2000(FileName)


# Tool: 137
@mcp.tool()
async def word_Document_ClosePrintPreview(this_Document_wordObjId: str):
	"""This tool calls the ClosePrintPreview methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ClosePrintPreview()


# Tool: 138
@mcp.tool()
async def word_Document_CheckConsistency(this_Document_wordObjId: str):
	"""This tool calls the CheckConsistency methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.CheckConsistency()


# Tool: 139
@mcp.tool()
async def word_Document_CreateLetterContent(this_Document_wordObjId: str, DateFormat: str, IncludeHeaderFooter: bool, PageDesign: str, LetterStyle: int, Letterhead: bool, LetterheadLocation: int, LetterheadSize: float, RecipientName: str, RecipientAddress: str, Salutation: str, SalutationType: int, RecipientReference: str, MailingInstructions: str, AttentionLine: str, Subject: str, CCList: str, ReturnAddress: str, SenderName: str, Closing: str, SenderCompany: str, SenderJobTitle: str, SenderInitials: str, EnclosureNumber: int, InfoBlock, RecipientCode, RecipientGender, ReturnAddressShortForm, SenderCity, SenderCode, SenderGender, SenderReference):
	"""This tool calls the CreateLetterContent methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		DateFormat: the DateFormat as str
		IncludeHeaderFooter: the IncludeHeaderFooter as bool
		PageDesign: the PageDesign as str
		LetterStyle: the LetterStyle as WdLetterStyle
		Letterhead: the Letterhead as bool
		LetterheadLocation: the LetterheadLocation as WdLetterheadLocation
		LetterheadSize: the LetterheadSize as float
		RecipientName: the RecipientName as str
		RecipientAddress: the RecipientAddress as str
		Salutation: the Salutation as str
		SalutationType: the SalutationType as WdSalutationType
		RecipientReference: the RecipientReference as str
		MailingInstructions: the MailingInstructions as str
		AttentionLine: the AttentionLine as str
		Subject: the Subject as str
		CCList: the CCList as str
		ReturnAddress: the ReturnAddress as str
		SenderName: the SenderName as str
		Closing: the Closing as str
		SenderCompany: the SenderCompany as str
		SenderJobTitle: the SenderJobTitle as str
		SenderInitials: the SenderInitials as str
		EnclosureNumber: the EnclosureNumber as int
		InfoBlock: the InfoBlock as VT_VARIANT
		RecipientCode: the RecipientCode as VT_VARIANT
		RecipientGender: the RecipientGender as VT_VARIANT
		ReturnAddressShortForm: the ReturnAddressShortForm as VT_VARIANT
		SenderCity: the SenderCity as VT_VARIANT
		SenderCode: the SenderCode as VT_VARIANT
		SenderGender: the SenderGender as VT_VARIANT
		SenderReference: the SenderReference as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	InfoBlock = tryParseString(InfoBlock)
	RecipientCode = tryParseString(RecipientCode)
	RecipientGender = tryParseString(RecipientGender)
	ReturnAddressShortForm = tryParseString(ReturnAddressShortForm)
	SenderCity = tryParseString(SenderCity)
	SenderCode = tryParseString(SenderCode)
	SenderGender = tryParseString(SenderGender)
	SenderReference = tryParseString(SenderReference)
	retVal = this_Document.CreateLetterContent(DateFormat, IncludeHeaderFooter, PageDesign, LetterStyle, Letterhead, LetterheadLocation, LetterheadSize, RecipientName, RecipientAddress, Salutation, SalutationType, RecipientReference, MailingInstructions, AttentionLine, Subject, CCList, ReturnAddress, SenderName, Closing, SenderCompany, SenderJobTitle, SenderInitials, EnclosureNumber, InfoBlock, RecipientCode, RecipientGender, ReturnAddressShortForm, SenderCity, SenderCode, SenderGender, SenderReference)
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "LetterContent"}


# Tool: 140
@mcp.tool()
async def word_Document_AcceptAllRevisions(this_Document_wordObjId: str):
	"""This tool calls the AcceptAllRevisions methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.AcceptAllRevisions()


# Tool: 141
@mcp.tool()
async def word_Document_RejectAllRevisions(this_Document_wordObjId: str):
	"""This tool calls the RejectAllRevisions methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.RejectAllRevisions()


# Tool: 142
@mcp.tool()
async def word_Document_DetectLanguage(this_Document_wordObjId: str):
	"""This tool calls the DetectLanguage methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.DetectLanguage()


# Tool: 143
@mcp.tool()
async def word_Document_ApplyTheme(this_Document_wordObjId: str, Name: str):
	"""This tool calls the ApplyTheme methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Name: the Name as str
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ApplyTheme(Name)


# Tool: 144
@mcp.tool()
async def word_Document_RemoveTheme(this_Document_wordObjId: str):
	"""This tool calls the RemoveTheme methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.RemoveTheme()


# Tool: 145
@mcp.tool()
async def word_Document_WebPagePreview(this_Document_wordObjId: str):
	"""This tool calls the WebPagePreview methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.WebPagePreview()


# Tool: 146
@mcp.tool()
async def word_Document_ReloadAs(this_Document_wordObjId: str, Encoding: int):
	"""This tool calls the ReloadAs methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Encoding: the Encoding as MsoEncoding
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ReloadAs(Encoding)


# Tool: 147
@mcp.tool()
async def word_Document_PrintOut2000(this_Document_wordObjId: str, Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight):
	"""This tool calls the PrintOut2000 methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Background: the Background as VT_VARIANT
		Append: the Append as VT_VARIANT
		Range: the Range as VT_VARIANT
		OutputFileName: the OutputFileName as VT_VARIANT
		From: the From as VT_VARIANT
		To: the To as VT_VARIANT
		Item: the Item as VT_VARIANT
		Copies: the Copies as VT_VARIANT
		Pages: the Pages as VT_VARIANT
		PageType: the PageType as VT_VARIANT
		PrintToFile: the PrintToFile as VT_VARIANT
		Collate: the Collate as VT_VARIANT
		ActivePrinterMacGX: the ActivePrinterMacGX as VT_VARIANT
		ManualDuplexPrint: the ManualDuplexPrint as VT_VARIANT
		PrintZoomColumn: the PrintZoomColumn as VT_VARIANT
		PrintZoomRow: the PrintZoomRow as VT_VARIANT
		PrintZoomPaperWidth: the PrintZoomPaperWidth as VT_VARIANT
		PrintZoomPaperHeight: the PrintZoomPaperHeight as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Background = tryParseString(Background)
	Append = tryParseString(Append)
	Range = tryParseString(Range)
	OutputFileName = tryParseString(OutputFileName)
	From = tryParseString(From)
	To = tryParseString(To)
	Item = tryParseString(Item)
	Copies = tryParseString(Copies)
	Pages = tryParseString(Pages)
	PageType = tryParseString(PageType)
	PrintToFile = tryParseString(PrintToFile)
	Collate = tryParseString(Collate)
	ActivePrinterMacGX = tryParseString(ActivePrinterMacGX)
	ManualDuplexPrint = tryParseString(ManualDuplexPrint)
	PrintZoomColumn = tryParseString(PrintZoomColumn)
	PrintZoomRow = tryParseString(PrintZoomRow)
	PrintZoomPaperWidth = tryParseString(PrintZoomPaperWidth)
	PrintZoomPaperHeight = tryParseString(PrintZoomPaperHeight)
	this_Document.PrintOut2000(Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight)


# Tool: 148
@mcp.tool()
async def word_Document_sblt(this_Document_wordObjId: str, s: str):
	"""This tool calls the sblt methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		s: the s as str
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.sblt(s)


# Tool: 149
@mcp.tool()
async def word_Document_ConvertVietDoc(this_Document_wordObjId: str, CodePageOrigin: int):
	"""This tool calls the ConvertVietDoc methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		CodePageOrigin: the CodePageOrigin as int
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ConvertVietDoc(CodePageOrigin)


# Tool: 150
@mcp.tool()
async def word_Document_PrintOut(this_Document_wordObjId: str, Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight):
	"""This tool calls the PrintOut methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Background: the Background as VT_VARIANT
		Append: the Append as VT_VARIANT
		Range: the Range as VT_VARIANT
		OutputFileName: the OutputFileName as VT_VARIANT
		From: the From as VT_VARIANT
		To: the To as VT_VARIANT
		Item: the Item as VT_VARIANT
		Copies: the Copies as VT_VARIANT
		Pages: the Pages as VT_VARIANT
		PageType: the PageType as VT_VARIANT
		PrintToFile: the PrintToFile as VT_VARIANT
		Collate: the Collate as VT_VARIANT
		ActivePrinterMacGX: the ActivePrinterMacGX as VT_VARIANT
		ManualDuplexPrint: the ManualDuplexPrint as VT_VARIANT
		PrintZoomColumn: the PrintZoomColumn as VT_VARIANT
		PrintZoomRow: the PrintZoomRow as VT_VARIANT
		PrintZoomPaperWidth: the PrintZoomPaperWidth as VT_VARIANT
		PrintZoomPaperHeight: the PrintZoomPaperHeight as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Background = tryParseString(Background)
	Append = tryParseString(Append)
	Range = tryParseString(Range)
	OutputFileName = tryParseString(OutputFileName)
	From = tryParseString(From)
	To = tryParseString(To)
	Item = tryParseString(Item)
	Copies = tryParseString(Copies)
	Pages = tryParseString(Pages)
	PageType = tryParseString(PageType)
	PrintToFile = tryParseString(PrintToFile)
	Collate = tryParseString(Collate)
	ActivePrinterMacGX = tryParseString(ActivePrinterMacGX)
	ManualDuplexPrint = tryParseString(ManualDuplexPrint)
	PrintZoomColumn = tryParseString(PrintZoomColumn)
	PrintZoomRow = tryParseString(PrintZoomRow)
	PrintZoomPaperWidth = tryParseString(PrintZoomPaperWidth)
	PrintZoomPaperHeight = tryParseString(PrintZoomPaperHeight)
	this_Document.PrintOut(Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight)


# Tool: 151
@mcp.tool()
async def word_Document_Compare2002(this_Document_wordObjId: str, Name: str, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles):
	"""This tool calls the Compare2002 methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Name: the Name as str
		AuthorName: the AuthorName as VT_VARIANT
		CompareTarget: the CompareTarget as VT_VARIANT
		DetectFormatChanges: the DetectFormatChanges as VT_VARIANT
		IgnoreAllComparisonWarnings: the IgnoreAllComparisonWarnings as VT_VARIANT
		AddToRecentFiles: the AddToRecentFiles as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	AuthorName = tryParseString(AuthorName)
	CompareTarget = tryParseString(CompareTarget)
	DetectFormatChanges = tryParseString(DetectFormatChanges)
	IgnoreAllComparisonWarnings = tryParseString(IgnoreAllComparisonWarnings)
	AddToRecentFiles = tryParseString(AddToRecentFiles)
	this_Document.Compare2002(Name, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles)


# Tool: 152
@mcp.tool()
async def word_Document_CheckIn(this_Document_wordObjId: str, SaveChanges: bool, Comments, MakePublic: bool):
	"""This tool calls the CheckIn methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		SaveChanges: the SaveChanges as bool
		Comments: the Comments as VT_VARIANT
		MakePublic: the MakePublic as bool
	"""
	this_Document = get_object(this_Document_wordObjId)
	Comments = tryParseString(Comments)
	this_Document.CheckIn(SaveChanges, Comments, MakePublic)


# Tool: 153
@mcp.tool()
async def word_Document_CanCheckin(this_Document_wordObjId: str):
	"""This tool calls the CanCheckin methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	retVal = this_Document.CanCheckin()
	return retVal


# Tool: 154
@mcp.tool()
async def word_Document_Merge(this_Document_wordObjId: str, FileName: str, MergeTarget, DetectFormatChanges, UseFormattingFrom, AddToRecentFiles):
	"""This tool calls the Merge methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as str
		MergeTarget: the MergeTarget as VT_VARIANT
		DetectFormatChanges: the DetectFormatChanges as VT_VARIANT
		UseFormattingFrom: the UseFormattingFrom as VT_VARIANT
		AddToRecentFiles: the AddToRecentFiles as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	MergeTarget = tryParseString(MergeTarget)
	DetectFormatChanges = tryParseString(DetectFormatChanges)
	UseFormattingFrom = tryParseString(UseFormattingFrom)
	AddToRecentFiles = tryParseString(AddToRecentFiles)
	this_Document.Merge(FileName, MergeTarget, DetectFormatChanges, UseFormattingFrom, AddToRecentFiles)


# Tool: 155
@mcp.tool()
async def word_Document_SendForReview(this_Document_wordObjId: str, Recipients, Subject, ShowMessage, IncludeAttachment):
	"""This tool calls the SendForReview methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Recipients: the Recipients as VT_VARIANT
		Subject: the Subject as VT_VARIANT
		ShowMessage: the ShowMessage as VT_VARIANT
		IncludeAttachment: the IncludeAttachment as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Recipients = tryParseString(Recipients)
	Subject = tryParseString(Subject)
	ShowMessage = tryParseString(ShowMessage)
	IncludeAttachment = tryParseString(IncludeAttachment)
	this_Document.SendForReview(Recipients, Subject, ShowMessage, IncludeAttachment)


# Tool: 156
@mcp.tool()
async def word_Document_ReplyWithChanges(this_Document_wordObjId: str, ShowMessage):
	"""This tool calls the ReplyWithChanges methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		ShowMessage: the ShowMessage as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	ShowMessage = tryParseString(ShowMessage)
	this_Document.ReplyWithChanges(ShowMessage)


# Tool: 157
@mcp.tool()
async def word_Document_EndReview(this_Document_wordObjId: str):
	"""This tool calls the EndReview methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.EndReview()


# Tool: 158
@mcp.tool()
async def word_Document_SetPasswordEncryptionOptions(this_Document_wordObjId: str, PasswordEncryptionProvider: str, PasswordEncryptionAlgorithm: str, PasswordEncryptionKeyLength: int, PasswordEncryptionFileProperties):
	"""This tool calls the SetPasswordEncryptionOptions methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		PasswordEncryptionProvider: the PasswordEncryptionProvider as str
		PasswordEncryptionAlgorithm: the PasswordEncryptionAlgorithm as str
		PasswordEncryptionKeyLength: the PasswordEncryptionKeyLength as int
		PasswordEncryptionFileProperties: the PasswordEncryptionFileProperties as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	PasswordEncryptionFileProperties = tryParseString(PasswordEncryptionFileProperties)
	this_Document.SetPasswordEncryptionOptions(PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties)


# Tool: 159
@mcp.tool()
async def word_Document_RecheckSmartTags(this_Document_wordObjId: str):
	"""This tool calls the RecheckSmartTags methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.RecheckSmartTags()


# Tool: 160
@mcp.tool()
async def word_Document_RemoveSmartTags(this_Document_wordObjId: str):
	"""This tool calls the RemoveSmartTags methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.RemoveSmartTags()


# Tool: 161
@mcp.tool()
async def word_Document_SetDefaultTableStyle(this_Document_wordObjId: str, Style, SetInTemplate: bool):
	"""This tool calls the SetDefaultTableStyle methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Style: the Style as VT_VARIANT
		SetInTemplate: the SetInTemplate as bool
	"""
	this_Document = get_object(this_Document_wordObjId)
	Style = tryParseString(Style)
	this_Document.SetDefaultTableStyle(Style, SetInTemplate)


# Tool: 162
@mcp.tool()
async def word_Document_DeleteAllComments(this_Document_wordObjId: str):
	"""This tool calls the DeleteAllComments methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.DeleteAllComments()


# Tool: 163
@mcp.tool()
async def word_Document_AcceptAllRevisionsShown(this_Document_wordObjId: str):
	"""This tool calls the AcceptAllRevisionsShown methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.AcceptAllRevisionsShown()


# Tool: 164
@mcp.tool()
async def word_Document_RejectAllRevisionsShown(this_Document_wordObjId: str):
	"""This tool calls the RejectAllRevisionsShown methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.RejectAllRevisionsShown()


# Tool: 165
@mcp.tool()
async def word_Document_DeleteAllCommentsShown(this_Document_wordObjId: str):
	"""This tool calls the DeleteAllCommentsShown methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.DeleteAllCommentsShown()


# Tool: 166
@mcp.tool()
async def word_Document_ResetFormFields(this_Document_wordObjId: str):
	"""This tool calls the ResetFormFields methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ResetFormFields()


# Tool: 167
@mcp.tool()
async def word_Document_SaveAs(this_Document_wordObjId: str, FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks):
	"""This tool calls the SaveAs methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as VT_VARIANT
		FileFormat: the FileFormat as VT_VARIANT
		LockComments: the LockComments as VT_VARIANT
		Password: the Password as VT_VARIANT
		AddToRecentFiles: the AddToRecentFiles as VT_VARIANT
		WritePassword: the WritePassword as VT_VARIANT
		ReadOnlyRecommended: the ReadOnlyRecommended as VT_VARIANT
		EmbedTrueTypeFonts: the EmbedTrueTypeFonts as VT_VARIANT
		SaveNativePictureFormat: the SaveNativePictureFormat as VT_VARIANT
		SaveFormsData: the SaveFormsData as VT_VARIANT
		SaveAsAOCELetter: the SaveAsAOCELetter as VT_VARIANT
		Encoding: the Encoding as VT_VARIANT
		InsertLineBreaks: the InsertLineBreaks as VT_VARIANT
		AllowSubstitutions: the AllowSubstitutions as VT_VARIANT
		LineEnding: the LineEnding as VT_VARIANT
		AddBiDiMarks: the AddBiDiMarks as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	FileName = tryParseString(FileName)
	FileFormat = tryParseString(FileFormat)
	LockComments = tryParseString(LockComments)
	Password = tryParseString(Password)
	AddToRecentFiles = tryParseString(AddToRecentFiles)
	WritePassword = tryParseString(WritePassword)
	ReadOnlyRecommended = tryParseString(ReadOnlyRecommended)
	EmbedTrueTypeFonts = tryParseString(EmbedTrueTypeFonts)
	SaveNativePictureFormat = tryParseString(SaveNativePictureFormat)
	SaveFormsData = tryParseString(SaveFormsData)
	SaveAsAOCELetter = tryParseString(SaveAsAOCELetter)
	Encoding = tryParseString(Encoding)
	InsertLineBreaks = tryParseString(InsertLineBreaks)
	AllowSubstitutions = tryParseString(AllowSubstitutions)
	LineEnding = tryParseString(LineEnding)
	AddBiDiMarks = tryParseString(AddBiDiMarks)
	this_Document.SaveAs(FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks)


# Tool: 168
@mcp.tool()
async def word_Document_CheckNewSmartTags(this_Document_wordObjId: str):
	"""This tool calls the CheckNewSmartTags methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.CheckNewSmartTags()


# Tool: 169
@mcp.tool()
async def word_Document_SendFaxOverInternet(this_Document_wordObjId: str, Recipients, Subject, ShowMessage):
	"""This tool calls the SendFaxOverInternet methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Recipients: the Recipients as VT_VARIANT
		Subject: the Subject as VT_VARIANT
		ShowMessage: the ShowMessage as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Recipients = tryParseString(Recipients)
	Subject = tryParseString(Subject)
	ShowMessage = tryParseString(ShowMessage)
	this_Document.SendFaxOverInternet(Recipients, Subject, ShowMessage)


# Tool: 170
@mcp.tool()
async def word_Document_TransformDocument(this_Document_wordObjId: str, Path: str, DataOnly: bool):
	"""This tool calls the TransformDocument methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Path: the Path as str
		DataOnly: the DataOnly as bool
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.TransformDocument(Path, DataOnly)


# Tool: 171
@mcp.tool()
async def word_Document_Protect(this_Document_wordObjId: str, Type: int, NoReset, Password, UseIRM, EnforceStyleLock):
	"""This tool calls the Protect methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Type: the Type as WdProtectionType
		NoReset: the NoReset as VT_VARIANT
		Password: the Password as VT_VARIANT
		UseIRM: the UseIRM as VT_VARIANT
		EnforceStyleLock: the EnforceStyleLock as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	NoReset = tryParseString(NoReset)
	Password = tryParseString(Password)
	UseIRM = tryParseString(UseIRM)
	EnforceStyleLock = tryParseString(EnforceStyleLock)
	this_Document.Protect(Type, NoReset, Password, UseIRM, EnforceStyleLock)


# Tool: 172
@mcp.tool()
async def word_Document_SelectAllEditableRanges(this_Document_wordObjId: str, EditorID):
	"""This tool calls the SelectAllEditableRanges methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		EditorID: the EditorID as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	EditorID = tryParseString(EditorID)
	this_Document.SelectAllEditableRanges(EditorID)


# Tool: 173
@mcp.tool()
async def word_Document_DeleteAllEditableRanges(this_Document_wordObjId: str, EditorID):
	"""This tool calls the DeleteAllEditableRanges methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		EditorID: the EditorID as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	EditorID = tryParseString(EditorID)
	this_Document.DeleteAllEditableRanges(EditorID)


# Tool: 174
@mcp.tool()
async def word_Document_DeleteAllInkAnnotations(this_Document_wordObjId: str):
	"""This tool calls the DeleteAllInkAnnotations methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.DeleteAllInkAnnotations()


# Tool: 175
@mcp.tool()
async def word_Document_AddDocumentWorkspaceHeader(this_Document_wordObjId: str, RichFormat: bool, Url: str, Title: str, Description: str, ID: str):
	"""This tool calls the AddDocumentWorkspaceHeader methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		RichFormat: the RichFormat as bool
		Url: the Url as str
		Title: the Title as str
		Description: the Description as str
		ID: the ID as str
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.AddDocumentWorkspaceHeader(RichFormat, Url, Title, Description, ID)


# Tool: 176
@mcp.tool()
async def word_Document_RemoveDocumentWorkspaceHeader(this_Document_wordObjId: str, ID: str):
	"""This tool calls the RemoveDocumentWorkspaceHeader methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		ID: the ID as str
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.RemoveDocumentWorkspaceHeader(ID)


# Tool: 177
@mcp.tool()
async def word_Document_Compare(this_Document_wordObjId: str, Name: str, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles, RemovePersonalInformation, RemoveDateAndTime):
	"""This tool calls the Compare methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Name: the Name as str
		AuthorName: the AuthorName as VT_VARIANT
		CompareTarget: the CompareTarget as VT_VARIANT
		DetectFormatChanges: the DetectFormatChanges as VT_VARIANT
		IgnoreAllComparisonWarnings: the IgnoreAllComparisonWarnings as VT_VARIANT
		AddToRecentFiles: the AddToRecentFiles as VT_VARIANT
		RemovePersonalInformation: the RemovePersonalInformation as VT_VARIANT
		RemoveDateAndTime: the RemoveDateAndTime as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	AuthorName = tryParseString(AuthorName)
	CompareTarget = tryParseString(CompareTarget)
	DetectFormatChanges = tryParseString(DetectFormatChanges)
	IgnoreAllComparisonWarnings = tryParseString(IgnoreAllComparisonWarnings)
	AddToRecentFiles = tryParseString(AddToRecentFiles)
	RemovePersonalInformation = tryParseString(RemovePersonalInformation)
	RemoveDateAndTime = tryParseString(RemoveDateAndTime)
	this_Document.Compare(Name, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles, RemovePersonalInformation, RemoveDateAndTime)


# Tool: 178
@mcp.tool()
async def word_Document_RemoveLockedStyles(this_Document_wordObjId: str):
	"""This tool calls the RemoveLockedStyles methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.RemoveLockedStyles()


# Tool: 179
@mcp.tool()
async def word_Document_SelectSingleNode(this_Document_wordObjId: str, XPath: str, PrefixMapping: str, FastSearchSkippingTextNodes: bool):
	"""This tool calls the SelectSingleNode methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		XPath: the XPath as str
		PrefixMapping: the PrefixMapping as str
		FastSearchSkippingTextNodes: the FastSearchSkippingTextNodes as bool
	"""
	this_Document = get_object(this_Document_wordObjId)
	retVal = this_Document.SelectSingleNode(XPath, PrefixMapping, FastSearchSkippingTextNodes)
	try:
		local_BaseName = retVal.BaseName
	except:
		local_BaseName = None
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_NamespaceURI = retVal.NamespaceURI
	except:
		local_NamespaceURI = None
	try:
		local_NodeType = retVal.NodeType
	except:
		local_NodeType = None
	try:
		local_NodeValue = retVal.NodeValue
	except:
		local_NodeValue = None
	try:
		local_HasChildNodes = retVal.HasChildNodes
	except:
		local_HasChildNodes = None
	try:
		local_Level = retVal.Level
	except:
		local_Level = None
	try:
		local_ValidationStatus = retVal.ValidationStatus
	except:
		local_ValidationStatus = None
	try:
		local_ValidationErrorText = retVal.ValidationErrorText
	except:
		local_ValidationErrorText = None
	try:
		local_PlaceholderText = retVal.PlaceholderText
	except:
		local_PlaceholderText = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "XMLNode", "BaseName": local_BaseName, "Text": local_Text, "NamespaceURI": local_NamespaceURI, "NodeType": local_NodeType, "NodeValue": local_NodeValue, "HasChildNodes": local_HasChildNodes, "Level": local_Level, "ValidationStatus": local_ValidationStatus, "ValidationErrorText": local_ValidationErrorText, "PlaceholderText": local_PlaceholderText, }


# Tool: 180
@mcp.tool()
async def word_Document_SelectNodes(this_Document_wordObjId: str, XPath: str, PrefixMapping: str, FastSearchSkippingTextNodes: bool):
	"""This tool calls the SelectNodes methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		XPath: the XPath as str
		PrefixMapping: the PrefixMapping as str
		FastSearchSkippingTextNodes: the FastSearchSkippingTextNodes as bool
	"""
	this_Document = get_object(this_Document_wordObjId)
	retVal = this_Document.SelectNodes(XPath, PrefixMapping, FastSearchSkippingTextNodes)
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "XMLNodes", "Count": local_Count, }


# Tool: 181
@mcp.tool()
async def word_Document_RemoveDocumentInformation(this_Document_wordObjId: str, RemoveDocInfoType: int):
	"""This tool calls the RemoveDocumentInformation methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		RemoveDocInfoType: the RemoveDocInfoType as WdRemoveDocInfoType
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.RemoveDocumentInformation(RemoveDocInfoType)


# Tool: 182
@mcp.tool()
async def word_Document_CheckInWithVersion(this_Document_wordObjId: str, SaveChanges: bool, Comments, MakePublic: bool, VersionType):
	"""This tool calls the CheckInWithVersion methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		SaveChanges: the SaveChanges as bool
		Comments: the Comments as VT_VARIANT
		MakePublic: the MakePublic as bool
		VersionType: the VersionType as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Comments = tryParseString(Comments)
	VersionType = tryParseString(VersionType)
	this_Document.CheckInWithVersion(SaveChanges, Comments, MakePublic, VersionType)


# Tool: 183
@mcp.tool()
async def word_Document_Dummy2(this_Document_wordObjId: str):
	"""This tool calls the Dummy2 methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.Dummy2()


# Tool: 184
@mcp.tool()
async def word_Document_LockServerFile(this_Document_wordObjId: str):
	"""This tool calls the LockServerFile methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.LockServerFile()


# Tool: 185
@mcp.tool()
async def word_Document_GetWorkflowTasks(this_Document_wordObjId: str):
	"""This tool calls the GetWorkflowTasks methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	retVal = this_Document.GetWorkflowTasks()
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "WorkflowTasks"}


# Tool: 186
@mcp.tool()
async def word_Document_GetWorkflowTemplates(this_Document_wordObjId: str):
	"""This tool calls the GetWorkflowTemplates methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	retVal = this_Document.GetWorkflowTemplates()
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "WorkflowTemplates"}


# Tool: 187
@mcp.tool()
async def word_Document_Dummy4(this_Document_wordObjId: str):
	"""This tool calls the Dummy4 methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.Dummy4()


# Tool: 188
@mcp.tool()
async def word_Document_AddMeetingWorkspaceHeader(this_Document_wordObjId: str, SkipIfAbsent: bool, Url: str, Title: str, Description: str, ID: str):
	"""This tool calls the AddMeetingWorkspaceHeader methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		SkipIfAbsent: the SkipIfAbsent as bool
		Url: the Url as str
		Title: the Title as str
		Description: the Description as str
		ID: the ID as str
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.AddMeetingWorkspaceHeader(SkipIfAbsent, Url, Title, Description, ID)


# Tool: 189
@mcp.tool()
async def word_Document_SaveAsQuickStyleSet(this_Document_wordObjId: str, FileName: str):
	"""This tool calls the SaveAsQuickStyleSet methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as str
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.SaveAsQuickStyleSet(FileName)


# Tool: 190
@mcp.tool()
async def word_Document_ApplyQuickStyleSet(this_Document_wordObjId: str, Name: str):
	"""This tool calls the ApplyQuickStyleSet methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Name: the Name as str
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ApplyQuickStyleSet(Name)


# Tool: 191
@mcp.tool()
async def word_Document_ApplyDocumentTheme(this_Document_wordObjId: str, FileName: str):
	"""This tool calls the ApplyDocumentTheme methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as str
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ApplyDocumentTheme(FileName)


# Tool: 192
@mcp.tool()
async def word_Document_SelectLinkedControls(this_Document_wordObjId: str, Node_wordObjId: str):
	"""This tool calls the SelectLinkedControls methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Node_wordObjId: 		To pass this object, send in the __WordObjectId of the CustomXMLNode object as was obtained from a previous return value
	"""
	this_Document = get_object(this_Document_wordObjId)
	Node = get_object(Node_wordObjId)
	retVal = this_Document.SelectLinkedControls(Node)
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ContentControls", "Count": local_Count, }


# Tool: 193
@mcp.tool()
async def word_Document_SelectUnlinkedControls(this_Document_wordObjId: str, Stream_wordObjId: str):
	"""This tool calls the SelectUnlinkedControls methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Stream_wordObjId: 		To pass this object, send in the __WordObjectId of the CustomXMLPart object as was obtained from a previous return value
	"""
	this_Document = get_object(this_Document_wordObjId)
	Stream = get_object(Stream_wordObjId)
	retVal = this_Document.SelectUnlinkedControls(Stream)
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ContentControls", "Count": local_Count, }


# Tool: 194
@mcp.tool()
async def word_Document_SelectContentControlsByTitle(this_Document_wordObjId: str, Title: str):
	"""This tool calls the SelectContentControlsByTitle methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Title: the Title as str
	"""
	this_Document = get_object(this_Document_wordObjId)
	retVal = this_Document.SelectContentControlsByTitle(Title)
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ContentControls", "Count": local_Count, }


# Tool: 195
@mcp.tool()
async def word_Document_ExportAsFixedFormat(this_Document_wordObjId: str, OutputFileName: str, ExportFormat: int, OpenAfterExport: bool, OptimizeFor: int, Range: int, From: int, To: int, Item: int, IncludeDocProps: bool, KeepIRM: bool, CreateBookmarks: int, DocStructureTags: bool, BitmapMissingFonts: bool, UseISO19005_1: bool, FixedFormatExtClassPtr):
	"""This tool calls the ExportAsFixedFormat methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		OutputFileName: the OutputFileName as str
		ExportFormat: the ExportFormat as WdExportFormat
		OpenAfterExport: the OpenAfterExport as bool
		OptimizeFor: the OptimizeFor as WdExportOptimizeFor
		Range: the Range as WdExportRange
		From: the From as int
		To: the To as int
		Item: the Item as WdExportItem
		IncludeDocProps: the IncludeDocProps as bool
		KeepIRM: the KeepIRM as bool
		CreateBookmarks: the CreateBookmarks as WdExportCreateBookmarks
		DocStructureTags: the DocStructureTags as bool
		BitmapMissingFonts: the BitmapMissingFonts as bool
		UseISO19005_1: the UseISO19005_1 as bool
		FixedFormatExtClassPtr: the FixedFormatExtClassPtr as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	FixedFormatExtClassPtr = tryParseString(FixedFormatExtClassPtr)
	this_Document.ExportAsFixedFormat(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range, From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr)


# Tool: 196
@mcp.tool()
async def word_Document_FreezeLayout(this_Document_wordObjId: str):
	"""This tool calls the FreezeLayout methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.FreezeLayout()


# Tool: 197
@mcp.tool()
async def word_Document_UnfreezeLayout(this_Document_wordObjId: str):
	"""This tool calls the UnfreezeLayout methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.UnfreezeLayout()


# Tool: 198
@mcp.tool()
async def word_Document_DowngradeDocument(this_Document_wordObjId: str):
	"""This tool calls the DowngradeDocument methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.DowngradeDocument()


# Tool: 199
@mcp.tool()
async def word_Document_Convert(this_Document_wordObjId: str):
	"""This tool calls the Convert methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.Convert()


# Tool: 200
@mcp.tool()
async def word_Document_SelectContentControlsByTag(this_Document_wordObjId: str, Tag: str):
	"""This tool calls the SelectContentControlsByTag methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Tag: the Tag as str
	"""
	this_Document = get_object(this_Document_wordObjId)
	retVal = this_Document.SelectContentControlsByTag(Tag)
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ContentControls", "Count": local_Count, }


# Tool: 201
@mcp.tool()
async def word_Document_ConvertAutoHyphens(this_Document_wordObjId: str):
	"""This tool calls the ConvertAutoHyphens methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.ConvertAutoHyphens()


# Tool: 202
@mcp.tool()
async def word_Document_ApplyQuickStyleSet2(this_Document_wordObjId: str, Style):
	"""This tool calls the ApplyQuickStyleSet2 methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Style: the Style as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	Style = tryParseString(Style)
	this_Document.ApplyQuickStyleSet2(Style)


# Tool: 203
@mcp.tool()
async def word_Document_SaveAs2(this_Document_wordObjId: str, FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks, CompatibilityMode):
	"""This tool calls the SaveAs2 methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as VT_VARIANT
		FileFormat: the FileFormat as VT_VARIANT
		LockComments: the LockComments as VT_VARIANT
		Password: the Password as VT_VARIANT
		AddToRecentFiles: the AddToRecentFiles as VT_VARIANT
		WritePassword: the WritePassword as VT_VARIANT
		ReadOnlyRecommended: the ReadOnlyRecommended as VT_VARIANT
		EmbedTrueTypeFonts: the EmbedTrueTypeFonts as VT_VARIANT
		SaveNativePictureFormat: the SaveNativePictureFormat as VT_VARIANT
		SaveFormsData: the SaveFormsData as VT_VARIANT
		SaveAsAOCELetter: the SaveAsAOCELetter as VT_VARIANT
		Encoding: the Encoding as VT_VARIANT
		InsertLineBreaks: the InsertLineBreaks as VT_VARIANT
		AllowSubstitutions: the AllowSubstitutions as VT_VARIANT
		LineEnding: the LineEnding as VT_VARIANT
		AddBiDiMarks: the AddBiDiMarks as VT_VARIANT
		CompatibilityMode: the CompatibilityMode as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	FileName = tryParseString(FileName)
	FileFormat = tryParseString(FileFormat)
	LockComments = tryParseString(LockComments)
	Password = tryParseString(Password)
	AddToRecentFiles = tryParseString(AddToRecentFiles)
	WritePassword = tryParseString(WritePassword)
	ReadOnlyRecommended = tryParseString(ReadOnlyRecommended)
	EmbedTrueTypeFonts = tryParseString(EmbedTrueTypeFonts)
	SaveNativePictureFormat = tryParseString(SaveNativePictureFormat)
	SaveFormsData = tryParseString(SaveFormsData)
	SaveAsAOCELetter = tryParseString(SaveAsAOCELetter)
	Encoding = tryParseString(Encoding)
	InsertLineBreaks = tryParseString(InsertLineBreaks)
	AllowSubstitutions = tryParseString(AllowSubstitutions)
	LineEnding = tryParseString(LineEnding)
	AddBiDiMarks = tryParseString(AddBiDiMarks)
	CompatibilityMode = tryParseString(CompatibilityMode)
	this_Document.SaveAs2(FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks, CompatibilityMode)


# Tool: 204
@mcp.tool()
async def word_Document_SetCompatibilityMode(this_Document_wordObjId: str, Mode: int):
	"""This tool calls the SetCompatibilityMode methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		Mode: the Mode as int
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.SetCompatibilityMode(Mode)


# Tool: 205
@mcp.tool()
async def word_Document_ReturnToLastReadPosition(this_Document_wordObjId: str):
	"""This tool calls the ReturnToLastReadPosition methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
"""
	this_Document = get_object(this_Document_wordObjId)
	retVal = this_Document.ReturnToLastReadPosition()
	return retVal


# Tool: 206
@mcp.tool()
async def word_Document_SaveCopyAs(this_Document_wordObjId: str, FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks, CompatibilityMode):
	"""This tool calls the SaveCopyAs methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as VT_VARIANT
		FileFormat: the FileFormat as VT_VARIANT
		LockComments: the LockComments as VT_VARIANT
		Password: the Password as VT_VARIANT
		AddToRecentFiles: the AddToRecentFiles as VT_VARIANT
		WritePassword: the WritePassword as VT_VARIANT
		ReadOnlyRecommended: the ReadOnlyRecommended as VT_VARIANT
		EmbedTrueTypeFonts: the EmbedTrueTypeFonts as VT_VARIANT
		SaveNativePictureFormat: the SaveNativePictureFormat as VT_VARIANT
		SaveFormsData: the SaveFormsData as VT_VARIANT
		SaveAsAOCELetter: the SaveAsAOCELetter as VT_VARIANT
		Encoding: the Encoding as VT_VARIANT
		InsertLineBreaks: the InsertLineBreaks as VT_VARIANT
		AllowSubstitutions: the AllowSubstitutions as VT_VARIANT
		LineEnding: the LineEnding as VT_VARIANT
		AddBiDiMarks: the AddBiDiMarks as VT_VARIANT
		CompatibilityMode: the CompatibilityMode as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	FileName = tryParseString(FileName)
	FileFormat = tryParseString(FileFormat)
	LockComments = tryParseString(LockComments)
	Password = tryParseString(Password)
	AddToRecentFiles = tryParseString(AddToRecentFiles)
	WritePassword = tryParseString(WritePassword)
	ReadOnlyRecommended = tryParseString(ReadOnlyRecommended)
	EmbedTrueTypeFonts = tryParseString(EmbedTrueTypeFonts)
	SaveNativePictureFormat = tryParseString(SaveNativePictureFormat)
	SaveFormsData = tryParseString(SaveFormsData)
	SaveAsAOCELetter = tryParseString(SaveAsAOCELetter)
	Encoding = tryParseString(Encoding)
	InsertLineBreaks = tryParseString(InsertLineBreaks)
	AllowSubstitutions = tryParseString(AllowSubstitutions)
	LineEnding = tryParseString(LineEnding)
	AddBiDiMarks = tryParseString(AddBiDiMarks)
	CompatibilityMode = tryParseString(CompatibilityMode)
	this_Document.SaveCopyAs(FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks, CompatibilityMode)


# Tool: 207
@mcp.tool()
async def word_Document_InsertAppForOfficeTest(this_Document_wordObjId: str, ID: str, Type: int, Version: str, StoreType: int, StoreId: str, AssetId: str, AssetStoreId: str, Width: int, Height: int):
	"""This tool calls the InsertAppForOfficeTest methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		ID: the ID as str
		Type: the Type as int
		Version: the Version as str
		StoreType: the StoreType as int
		StoreId: the StoreId as str
		AssetId: the AssetId as str
		AssetStoreId: the AssetStoreId as str
		Width: the Width as int
		Height: the Height as int
	"""
	this_Document = get_object(this_Document_wordObjId)
	this_Document.InsertAppForOfficeTest(ID, Type, Version, StoreType, StoreId, AssetId, AssetStoreId, Width, Height)


# Tool: 208
@mcp.tool()
async def word_Document_ExportAsFixedFormat2(this_Document_wordObjId: str, OutputFileName: str, ExportFormat: int, OpenAfterExport: bool, OptimizeFor: int, Range: int, From: int, To: int, Item: int, IncludeDocProps: bool, KeepIRM: bool, CreateBookmarks: int, DocStructureTags: bool, BitmapMissingFonts: bool, UseISO19005_1: bool, OptimizeForImageQuality: bool, FixedFormatExtClassPtr):
	"""This tool calls the ExportAsFixedFormat2 methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		OutputFileName: the OutputFileName as str
		ExportFormat: the ExportFormat as WdExportFormat
		OpenAfterExport: the OpenAfterExport as bool
		OptimizeFor: the OptimizeFor as WdExportOptimizeFor
		Range: the Range as WdExportRange
		From: the From as int
		To: the To as int
		Item: the Item as WdExportItem
		IncludeDocProps: the IncludeDocProps as bool
		KeepIRM: the KeepIRM as bool
		CreateBookmarks: the CreateBookmarks as WdExportCreateBookmarks
		DocStructureTags: the DocStructureTags as bool
		BitmapMissingFonts: the BitmapMissingFonts as bool
		UseISO19005_1: the UseISO19005_1 as bool
		OptimizeForImageQuality: the OptimizeForImageQuality as bool
		FixedFormatExtClassPtr: the FixedFormatExtClassPtr as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	FixedFormatExtClassPtr = tryParseString(FixedFormatExtClassPtr)
	this_Document.ExportAsFixedFormat2(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range, From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, OptimizeForImageQuality, FixedFormatExtClassPtr)


# Tool: 209
@mcp.tool()
async def word_Document_ExportAsFixedFormat3(this_Document_wordObjId: str, OutputFileName: str, ExportFormat: int, OpenAfterExport: bool, OptimizeFor: int, Range: int, From: int, To: int, Item: int, IncludeDocProps: bool, KeepIRM: bool, CreateBookmarks: int, DocStructureTags: bool, BitmapMissingFonts: bool, UseISO19005_1: bool, OptimizeForImageQuality: bool, ImproveExportTagging: bool, FixedFormatExtClassPtr):
	"""This tool calls the ExportAsFixedFormat3 methodon an Document object. Pass the __WordObjectId of Document of the object you want to call the method on as the first parameter
	
	Parameters:
		OutputFileName: the OutputFileName as str
		ExportFormat: the ExportFormat as WdExportFormat
		OpenAfterExport: the OpenAfterExport as bool
		OptimizeFor: the OptimizeFor as WdExportOptimizeFor
		Range: the Range as WdExportRange
		From: the From as int
		To: the To as int
		Item: the Item as WdExportItem
		IncludeDocProps: the IncludeDocProps as bool
		KeepIRM: the KeepIRM as bool
		CreateBookmarks: the CreateBookmarks as WdExportCreateBookmarks
		DocStructureTags: the DocStructureTags as bool
		BitmapMissingFonts: the BitmapMissingFonts as bool
		UseISO19005_1: the UseISO19005_1 as bool
		OptimizeForImageQuality: the OptimizeForImageQuality as bool
		ImproveExportTagging: the ImproveExportTagging as bool
		FixedFormatExtClassPtr: the FixedFormatExtClassPtr as VT_VARIANT
	"""
	this_Document = get_object(this_Document_wordObjId)
	FixedFormatExtClassPtr = tryParseString(FixedFormatExtClassPtr)
	this_Document.ExportAsFixedFormat3(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range, From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, OptimizeForImageQuality, ImproveExportTagging, FixedFormatExtClassPtr)


# Tool: 210
@mcp.tool()
async def word_Document_get_Property(this_Document_wordObjId: str, propertyName: str):
	"""Gets properties of Document
	
	propertyName: Name of the property. Can be one of ...
		Name, BuiltInDocumentProperties, CustomDocumentProperties, Path, Bookmarks, Tables, Footnotes, Endnotes, Comments, Type, AutoHyphenation, HyphenateCaps, HyphenationZone, ConsecutiveHyphensLimit, Sections, Paragraphs, Words, Sentences, Characters, Fields, FormFields, Styles, Frames, TablesOfFigures, Variables, MailMerge, Envelope, FullName, Revisions, TablesOfContents, TablesOfAuthorities, PageSetup, Windows, HasRoutingSlip, RoutingSlip, Routed, TablesOfAuthoritiesCategories, Indexes, Saved, Content, ActiveWindow, Kind, ReadOnly, Subdocuments, IsMasterDocument, DefaultTabStop, EmbedTrueTypeFonts, SaveFormsData, ReadOnlyRecommended, SaveSubsetFonts, StoryRanges, CommandBars, IsSubdocument, SaveFormat, ProtectionType, Hyperlinks, Shapes, ListTemplates, Lists, UpdateStylesOnOpen, AttachedTemplate, InlineShapes, Background, GrammarChecked, SpellingChecked, ShowGrammaticalErrors, ShowSpellingErrors, Versions, ShowSummary, SummaryViewMode, SummaryLength, PrintFractionalWidths, PrintPostScriptOverText, Container, PrintFormsData, ListParagraphs, HasPassword, WriteReserved, UserControl, HasMailer, Mailer, ReadabilityStatistics, GrammaticalErrors, SpellingErrors, VBProject, FormsDesign, _CodeName, CodeName, SnapToGrid, SnapToShapes, GridDistanceHorizontal, GridDistanceVertical, GridOriginHorizontal, GridOriginVertical, GridSpaceBetweenHorizontalLines, GridSpaceBetweenVerticalLines, GridOriginFromMargin, KerningByAlgorithm, JustificationMode, FarEastLineBreakLevel, NoLineBreakBefore, NoLineBreakAfter, TrackRevisions, PrintRevisions, ShowRevisions, ActiveTheme, ActiveThemeDisplayName, Email, Scripts, LanguageDetected, FarEastLineBreakLanguage, Frameset, ClickAndTypeParagraphStyle, HTMLProject, WebOptions, OpenEncoding, SaveEncoding, OptimizeForWord97, VBASigned, MailEnvelope, DisableFeatures, DoNotEmbedSystemFonts, Signatures, DefaultTargetFrame, HTMLDivisions, DisableFeaturesIntroducedAfter, RemovePersonalInformation, SmartTags, EmbedSmartTags, SmartTagsAsXMLProps, TextEncoding, TextLineEnding, StyleSheets, DefaultTableStyle, PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties, EmbedLinguisticData, FormattingShowFont, FormattingShowClear, FormattingShowParagraph, FormattingShowNumbering, FormattingShowFilter, Permission, XMLNodes, XMLSchemaReferences, SmartDocument, SharedWorkspace, Sync, EnforceStyle, AutoFormatOverride, XMLSaveDataOnly, XMLHideNamespaces, XMLShowAdvancedErrors, XMLUseXSLTWhenSaving, XMLSaveThroughXSLT, DocumentLibraryVersions, ReadingModeLayoutFrozen, RemoveDateAndTime, ChildNodeSuggestions, XMLSchemaViolations, ReadingLayoutSizeX, ReadingLayoutSizeY, StyleSortMethod, ContentTypeProperties, TrackMoves, TrackFormatting, OMaths, ServerPolicy, ContentControls, DocumentInspectors, Bibliography, LockTheme, LockQuickStyleSet, OriginalDocumentTitle, RevisedDocumentTitle, CustomXMLParts, FormattingShowNextLevel, FormattingShowUserStyleName, Research, Final, OMathBreakBin, OMathBreakSub, OMathJc, OMathLeftMargin, OMathRightMargin, OMathWrap, OMathIntSubSupLim, OMathNarySupSubLim, OMathSmallFrac, WordOpenXML, DocumentTheme, HasVBProject, OMathFontName, EncryptionProvider, UseMathDefaults, CurrentRsid, DocID, CompatibilityMode, CoAuthoring, Broadcast, ChartDataPointTrack, IsInAutosave, WorkIdentity, AutoSaveOn, SensitivityLabel, TrackJustMyRevisions
	"""
	this_Document = get_object(this_Document_wordObjId)
	
	EnsureWord()
	if (propertyName == "Name"):
		retVal = this_Document.Name
		return retVal
	if (propertyName == "BuiltInDocumentProperties"):
		retVal = this_Document.BuiltInDocumentProperties
		return retVal
	if (propertyName == "CustomDocumentProperties"):
		retVal = this_Document.CustomDocumentProperties
		return retVal
	if (propertyName == "Path"):
		retVal = this_Document.Path
		return retVal
	if (propertyName == "Bookmarks"):
		retVal = this_Document.Bookmarks
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_DefaultSorting = retVal.DefaultSorting
		except:
			local_DefaultSorting = None
		try:
			local_ShowHidden = retVal.ShowHidden
		except:
			local_ShowHidden = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Bookmarks", "Count": local_Count, "DefaultSorting": local_DefaultSorting, "ShowHidden": local_ShowHidden, }
	if (propertyName == "Tables"):
		retVal = this_Document.Tables
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Tables", "Count": local_Count, "NestingLevel": local_NestingLevel, }
	if (propertyName == "Footnotes"):
		retVal = this_Document.Footnotes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Location = retVal.Location
		except:
			local_Location = None
		try:
			local_NumberStyle = retVal.NumberStyle
		except:
			local_NumberStyle = None
		try:
			local_StartingNumber = retVal.StartingNumber
		except:
			local_StartingNumber = None
		try:
			local_NumberingRule = retVal.NumberingRule
		except:
			local_NumberingRule = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Footnotes", "Count": local_Count, "Location": local_Location, "NumberStyle": local_NumberStyle, "StartingNumber": local_StartingNumber, "NumberingRule": local_NumberingRule, }
	if (propertyName == "Endnotes"):
		retVal = this_Document.Endnotes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Location = retVal.Location
		except:
			local_Location = None
		try:
			local_NumberStyle = retVal.NumberStyle
		except:
			local_NumberStyle = None
		try:
			local_StartingNumber = retVal.StartingNumber
		except:
			local_StartingNumber = None
		try:
			local_NumberingRule = retVal.NumberingRule
		except:
			local_NumberingRule = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Endnotes", "Count": local_Count, "Location": local_Location, "NumberStyle": local_NumberStyle, "StartingNumber": local_StartingNumber, "NumberingRule": local_NumberingRule, }
	if (propertyName == "Comments"):
		retVal = this_Document.Comments
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_ShowBy = retVal.ShowBy
		except:
			local_ShowBy = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Comments", "Count": local_Count, "ShowBy": local_ShowBy, }
	if (propertyName == "Type"):
		retVal = this_Document.Type
		return retVal
	if (propertyName == "AutoHyphenation"):
		retVal = this_Document.AutoHyphenation
		return retVal
	if (propertyName == "HyphenateCaps"):
		retVal = this_Document.HyphenateCaps
		return retVal
	if (propertyName == "HyphenationZone"):
		retVal = this_Document.HyphenationZone
		return retVal
	if (propertyName == "ConsecutiveHyphensLimit"):
		retVal = this_Document.ConsecutiveHyphensLimit
		return retVal
	if (propertyName == "Sections"):
		retVal = this_Document.Sections
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Sections", "Count": local_Count, }
	if (propertyName == "Paragraphs"):
		retVal = this_Document.Paragraphs
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_Alignment = retVal.Alignment
		except:
			local_Alignment = None
		try:
			local_KeepTogether = retVal.KeepTogether
		except:
			local_KeepTogether = None
		try:
			local_KeepWithNext = retVal.KeepWithNext
		except:
			local_KeepWithNext = None
		try:
			local_PageBreakBefore = retVal.PageBreakBefore
		except:
			local_PageBreakBefore = None
		try:
			local_NoLineNumber = retVal.NoLineNumber
		except:
			local_NoLineNumber = None
		try:
			local_RightIndent = retVal.RightIndent
		except:
			local_RightIndent = None
		try:
			local_LeftIndent = retVal.LeftIndent
		except:
			local_LeftIndent = None
		try:
			local_FirstLineIndent = retVal.FirstLineIndent
		except:
			local_FirstLineIndent = None
		try:
			local_LineSpacing = retVal.LineSpacing
		except:
			local_LineSpacing = None
		try:
			local_LineSpacingRule = retVal.LineSpacingRule
		except:
			local_LineSpacingRule = None
		try:
			local_SpaceBefore = retVal.SpaceBefore
		except:
			local_SpaceBefore = None
		try:
			local_SpaceAfter = retVal.SpaceAfter
		except:
			local_SpaceAfter = None
		try:
			local_Hyphenation = retVal.Hyphenation
		except:
			local_Hyphenation = None
		try:
			local_WidowControl = retVal.WidowControl
		except:
			local_WidowControl = None
		try:
			local_FarEastLineBreakControl = retVal.FarEastLineBreakControl
		except:
			local_FarEastLineBreakControl = None
		try:
			local_WordWrap = retVal.WordWrap
		except:
			local_WordWrap = None
		try:
			local_HangingPunctuation = retVal.HangingPunctuation
		except:
			local_HangingPunctuation = None
		try:
			local_HalfWidthPunctuationOnTopOfLine = retVal.HalfWidthPunctuationOnTopOfLine
		except:
			local_HalfWidthPunctuationOnTopOfLine = None
		try:
			local_AddSpaceBetweenFarEastAndAlpha = retVal.AddSpaceBetweenFarEastAndAlpha
		except:
			local_AddSpaceBetweenFarEastAndAlpha = None
		try:
			local_AddSpaceBetweenFarEastAndDigit = retVal.AddSpaceBetweenFarEastAndDigit
		except:
			local_AddSpaceBetweenFarEastAndDigit = None
		try:
			local_BaseLineAlignment = retVal.BaseLineAlignment
		except:
			local_BaseLineAlignment = None
		try:
			local_AutoAdjustRightIndent = retVal.AutoAdjustRightIndent
		except:
			local_AutoAdjustRightIndent = None
		try:
			local_DisableLineHeightGrid = retVal.DisableLineHeightGrid
		except:
			local_DisableLineHeightGrid = None
		try:
			local_OutlineLevel = retVal.OutlineLevel
		except:
			local_OutlineLevel = None
		try:
			local_CharacterUnitRightIndent = retVal.CharacterUnitRightIndent
		except:
			local_CharacterUnitRightIndent = None
		try:
			local_CharacterUnitLeftIndent = retVal.CharacterUnitLeftIndent
		except:
			local_CharacterUnitLeftIndent = None
		try:
			local_CharacterUnitFirstLineIndent = retVal.CharacterUnitFirstLineIndent
		except:
			local_CharacterUnitFirstLineIndent = None
		try:
			local_LineUnitBefore = retVal.LineUnitBefore
		except:
			local_LineUnitBefore = None
		try:
			local_LineUnitAfter = retVal.LineUnitAfter
		except:
			local_LineUnitAfter = None
		try:
			local_ReadingOrder = retVal.ReadingOrder
		except:
			local_ReadingOrder = None
		try:
			local_SpaceBeforeAuto = retVal.SpaceBeforeAuto
		except:
			local_SpaceBeforeAuto = None
		try:
			local_SpaceAfterAuto = retVal.SpaceAfterAuto
		except:
			local_SpaceAfterAuto = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Paragraphs", "Count": local_Count, "Style": local_Style, "Alignment": local_Alignment, "KeepTogether": local_KeepTogether, "KeepWithNext": local_KeepWithNext, "PageBreakBefore": local_PageBreakBefore, "NoLineNumber": local_NoLineNumber, "RightIndent": local_RightIndent, "LeftIndent": local_LeftIndent, "FirstLineIndent": local_FirstLineIndent, "LineSpacing": local_LineSpacing, "LineSpacingRule": local_LineSpacingRule, "SpaceBefore": local_SpaceBefore, "SpaceAfter": local_SpaceAfter, "Hyphenation": local_Hyphenation, "WidowControl": local_WidowControl, "FarEastLineBreakControl": local_FarEastLineBreakControl, "WordWrap": local_WordWrap, "HangingPunctuation": local_HangingPunctuation, "HalfWidthPunctuationOnTopOfLine": local_HalfWidthPunctuationOnTopOfLine, "AddSpaceBetweenFarEastAndAlpha": local_AddSpaceBetweenFarEastAndAlpha, "AddSpaceBetweenFarEastAndDigit": local_AddSpaceBetweenFarEastAndDigit, "BaseLineAlignment": local_BaseLineAlignment, "AutoAdjustRightIndent": local_AutoAdjustRightIndent, "DisableLineHeightGrid": local_DisableLineHeightGrid, "OutlineLevel": local_OutlineLevel, "CharacterUnitRightIndent": local_CharacterUnitRightIndent, "CharacterUnitLeftIndent": local_CharacterUnitLeftIndent, "CharacterUnitFirstLineIndent": local_CharacterUnitFirstLineIndent, "LineUnitBefore": local_LineUnitBefore, "LineUnitAfter": local_LineUnitAfter, "ReadingOrder": local_ReadingOrder, "SpaceBeforeAuto": local_SpaceBeforeAuto, "SpaceAfterAuto": local_SpaceAfterAuto, }
	if (propertyName == "Words"):
		retVal = this_Document.Words
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Words", "Count": local_Count, }
	if (propertyName == "Sentences"):
		retVal = this_Document.Sentences
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Sentences", "Count": local_Count, }
	if (propertyName == "Characters"):
		retVal = this_Document.Characters
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Characters", "Count": local_Count, }
	if (propertyName == "Fields"):
		retVal = this_Document.Fields
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Locked = retVal.Locked
		except:
			local_Locked = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Fields", "Count": local_Count, "Locked": local_Locked, }
	if (propertyName == "FormFields"):
		retVal = this_Document.FormFields
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Shaded = retVal.Shaded
		except:
			local_Shaded = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "FormFields", "Count": local_Count, "Shaded": local_Shaded, }
	if (propertyName == "Styles"):
		retVal = this_Document.Styles
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Styles", "Count": local_Count, }
	if (propertyName == "Frames"):
		retVal = this_Document.Frames
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Frames", "Count": local_Count, }
	if (propertyName == "TablesOfFigures"):
		retVal = this_Document.TablesOfFigures
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Format = retVal.Format
		except:
			local_Format = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "TablesOfFigures", "Count": local_Count, "Format": local_Format, }
	if (propertyName == "Variables"):
		retVal = this_Document.Variables
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Variables", "Count": local_Count, }
	if (propertyName == "MailMerge"):
		retVal = this_Document.MailMerge
		try:
			local_MainDocumentType = retVal.MainDocumentType
		except:
			local_MainDocumentType = None
		try:
			local_State = retVal.State
		except:
			local_State = None
		try:
			local_Destination = retVal.Destination
		except:
			local_Destination = None
		try:
			local_ViewMailMergeFieldCodes = retVal.ViewMailMergeFieldCodes
		except:
			local_ViewMailMergeFieldCodes = None
		try:
			local_SuppressBlankLines = retVal.SuppressBlankLines
		except:
			local_SuppressBlankLines = None
		try:
			local_MailAsAttachment = retVal.MailAsAttachment
		except:
			local_MailAsAttachment = None
		try:
			local_MailAddressFieldName = retVal.MailAddressFieldName
		except:
			local_MailAddressFieldName = None
		try:
			local_MailSubject = retVal.MailSubject
		except:
			local_MailSubject = None
		try:
			local_HighlightMergeFields = retVal.HighlightMergeFields
		except:
			local_HighlightMergeFields = None
		try:
			local_MailFormat = retVal.MailFormat
		except:
			local_MailFormat = None
		try:
			local_ShowSendToCustom = retVal.ShowSendToCustom
		except:
			local_ShowSendToCustom = None
		try:
			local_WizardState = retVal.WizardState
		except:
			local_WizardState = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "MailMerge", "MainDocumentType": local_MainDocumentType, "State": local_State, "Destination": local_Destination, "ViewMailMergeFieldCodes": local_ViewMailMergeFieldCodes, "SuppressBlankLines": local_SuppressBlankLines, "MailAsAttachment": local_MailAsAttachment, "MailAddressFieldName": local_MailAddressFieldName, "MailSubject": local_MailSubject, "HighlightMergeFields": local_HighlightMergeFields, "MailFormat": local_MailFormat, "ShowSendToCustom": local_ShowSendToCustom, "WizardState": local_WizardState, }
	if (propertyName == "Envelope"):
		retVal = this_Document.Envelope
		try:
			local_DefaultPrintBarCode = retVal.DefaultPrintBarCode
		except:
			local_DefaultPrintBarCode = None
		try:
			local_DefaultPrintFIMA = retVal.DefaultPrintFIMA
		except:
			local_DefaultPrintFIMA = None
		try:
			local_DefaultHeight = retVal.DefaultHeight
		except:
			local_DefaultHeight = None
		try:
			local_DefaultWidth = retVal.DefaultWidth
		except:
			local_DefaultWidth = None
		try:
			local_DefaultSize = retVal.DefaultSize
		except:
			local_DefaultSize = None
		try:
			local_DefaultOmitReturnAddress = retVal.DefaultOmitReturnAddress
		except:
			local_DefaultOmitReturnAddress = None
		try:
			local_FeedSource = retVal.FeedSource
		except:
			local_FeedSource = None
		try:
			local_AddressFromLeft = retVal.AddressFromLeft
		except:
			local_AddressFromLeft = None
		try:
			local_AddressFromTop = retVal.AddressFromTop
		except:
			local_AddressFromTop = None
		try:
			local_ReturnAddressFromLeft = retVal.ReturnAddressFromLeft
		except:
			local_ReturnAddressFromLeft = None
		try:
			local_ReturnAddressFromTop = retVal.ReturnAddressFromTop
		except:
			local_ReturnAddressFromTop = None
		try:
			local_DefaultOrientation = retVal.DefaultOrientation
		except:
			local_DefaultOrientation = None
		try:
			local_DefaultFaceUp = retVal.DefaultFaceUp
		except:
			local_DefaultFaceUp = None
		try:
			local_Vertical = retVal.Vertical
		except:
			local_Vertical = None
		try:
			local_RecipientNamefromLeft = retVal.RecipientNamefromLeft
		except:
			local_RecipientNamefromLeft = None
		try:
			local_RecipientNamefromTop = retVal.RecipientNamefromTop
		except:
			local_RecipientNamefromTop = None
		try:
			local_RecipientPostalfromLeft = retVal.RecipientPostalfromLeft
		except:
			local_RecipientPostalfromLeft = None
		try:
			local_RecipientPostalfromTop = retVal.RecipientPostalfromTop
		except:
			local_RecipientPostalfromTop = None
		try:
			local_SenderNamefromLeft = retVal.SenderNamefromLeft
		except:
			local_SenderNamefromLeft = None
		try:
			local_SenderNamefromTop = retVal.SenderNamefromTop
		except:
			local_SenderNamefromTop = None
		try:
			local_SenderPostalfromLeft = retVal.SenderPostalfromLeft
		except:
			local_SenderPostalfromLeft = None
		try:
			local_SenderPostalfromTop = retVal.SenderPostalfromTop
		except:
			local_SenderPostalfromTop = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Envelope", "DefaultPrintBarCode": local_DefaultPrintBarCode, "DefaultPrintFIMA": local_DefaultPrintFIMA, "DefaultHeight": local_DefaultHeight, "DefaultWidth": local_DefaultWidth, "DefaultSize": local_DefaultSize, "DefaultOmitReturnAddress": local_DefaultOmitReturnAddress, "FeedSource": local_FeedSource, "AddressFromLeft": local_AddressFromLeft, "AddressFromTop": local_AddressFromTop, "ReturnAddressFromLeft": local_ReturnAddressFromLeft, "ReturnAddressFromTop": local_ReturnAddressFromTop, "DefaultOrientation": local_DefaultOrientation, "DefaultFaceUp": local_DefaultFaceUp, "Vertical": local_Vertical, "RecipientNamefromLeft": local_RecipientNamefromLeft, "RecipientNamefromTop": local_RecipientNamefromTop, "RecipientPostalfromLeft": local_RecipientPostalfromLeft, "RecipientPostalfromTop": local_RecipientPostalfromTop, "SenderNamefromLeft": local_SenderNamefromLeft, "SenderNamefromTop": local_SenderNamefromTop, "SenderPostalfromLeft": local_SenderPostalfromLeft, "SenderPostalfromTop": local_SenderPostalfromTop, }
	if (propertyName == "FullName"):
		retVal = this_Document.FullName
		return retVal
	if (propertyName == "Revisions"):
		retVal = this_Document.Revisions
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Revisions", "Count": local_Count, }
	if (propertyName == "TablesOfContents"):
		retVal = this_Document.TablesOfContents
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Format = retVal.Format
		except:
			local_Format = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "TablesOfContents", "Count": local_Count, "Format": local_Format, }
	if (propertyName == "TablesOfAuthorities"):
		retVal = this_Document.TablesOfAuthorities
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Format = retVal.Format
		except:
			local_Format = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "TablesOfAuthorities", "Count": local_Count, "Format": local_Format, }
	if (propertyName == "PageSetup"):
		retVal = this_Document.PageSetup
		try:
			local_TopMargin = retVal.TopMargin
		except:
			local_TopMargin = None
		try:
			local_BottomMargin = retVal.BottomMargin
		except:
			local_BottomMargin = None
		try:
			local_LeftMargin = retVal.LeftMargin
		except:
			local_LeftMargin = None
		try:
			local_RightMargin = retVal.RightMargin
		except:
			local_RightMargin = None
		try:
			local_Gutter = retVal.Gutter
		except:
			local_Gutter = None
		try:
			local_PageWidth = retVal.PageWidth
		except:
			local_PageWidth = None
		try:
			local_PageHeight = retVal.PageHeight
		except:
			local_PageHeight = None
		try:
			local_Orientation = retVal.Orientation
		except:
			local_Orientation = None
		try:
			local_FirstPageTray = retVal.FirstPageTray
		except:
			local_FirstPageTray = None
		try:
			local_OtherPagesTray = retVal.OtherPagesTray
		except:
			local_OtherPagesTray = None
		try:
			local_VerticalAlignment = retVal.VerticalAlignment
		except:
			local_VerticalAlignment = None
		try:
			local_MirrorMargins = retVal.MirrorMargins
		except:
			local_MirrorMargins = None
		try:
			local_HeaderDistance = retVal.HeaderDistance
		except:
			local_HeaderDistance = None
		try:
			local_FooterDistance = retVal.FooterDistance
		except:
			local_FooterDistance = None
		try:
			local_SectionStart = retVal.SectionStart
		except:
			local_SectionStart = None
		try:
			local_OddAndEvenPagesHeaderFooter = retVal.OddAndEvenPagesHeaderFooter
		except:
			local_OddAndEvenPagesHeaderFooter = None
		try:
			local_DifferentFirstPageHeaderFooter = retVal.DifferentFirstPageHeaderFooter
		except:
			local_DifferentFirstPageHeaderFooter = None
		try:
			local_SuppressEndnotes = retVal.SuppressEndnotes
		except:
			local_SuppressEndnotes = None
		try:
			local_PaperSize = retVal.PaperSize
		except:
			local_PaperSize = None
		try:
			local_TwoPagesOnOne = retVal.TwoPagesOnOne
		except:
			local_TwoPagesOnOne = None
		try:
			local_GutterOnTop = retVal.GutterOnTop
		except:
			local_GutterOnTop = None
		try:
			local_CharsLine = retVal.CharsLine
		except:
			local_CharsLine = None
		try:
			local_LinesPage = retVal.LinesPage
		except:
			local_LinesPage = None
		try:
			local_ShowGrid = retVal.ShowGrid
		except:
			local_ShowGrid = None
		try:
			local_GutterStyle = retVal.GutterStyle
		except:
			local_GutterStyle = None
		try:
			local_SectionDirection = retVal.SectionDirection
		except:
			local_SectionDirection = None
		try:
			local_LayoutMode = retVal.LayoutMode
		except:
			local_LayoutMode = None
		try:
			local_GutterPos = retVal.GutterPos
		except:
			local_GutterPos = None
		try:
			local_BookFoldPrinting = retVal.BookFoldPrinting
		except:
			local_BookFoldPrinting = None
		try:
			local_BookFoldRevPrinting = retVal.BookFoldRevPrinting
		except:
			local_BookFoldRevPrinting = None
		try:
			local_BookFoldPrintingSheets = retVal.BookFoldPrintingSheets
		except:
			local_BookFoldPrintingSheets = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "PageSetup", "TopMargin": local_TopMargin, "BottomMargin": local_BottomMargin, "LeftMargin": local_LeftMargin, "RightMargin": local_RightMargin, "Gutter": local_Gutter, "PageWidth": local_PageWidth, "PageHeight": local_PageHeight, "Orientation": local_Orientation, "FirstPageTray": local_FirstPageTray, "OtherPagesTray": local_OtherPagesTray, "VerticalAlignment": local_VerticalAlignment, "MirrorMargins": local_MirrorMargins, "HeaderDistance": local_HeaderDistance, "FooterDistance": local_FooterDistance, "SectionStart": local_SectionStart, "OddAndEvenPagesHeaderFooter": local_OddAndEvenPagesHeaderFooter, "DifferentFirstPageHeaderFooter": local_DifferentFirstPageHeaderFooter, "SuppressEndnotes": local_SuppressEndnotes, "PaperSize": local_PaperSize, "TwoPagesOnOne": local_TwoPagesOnOne, "GutterOnTop": local_GutterOnTop, "CharsLine": local_CharsLine, "LinesPage": local_LinesPage, "ShowGrid": local_ShowGrid, "GutterStyle": local_GutterStyle, "SectionDirection": local_SectionDirection, "LayoutMode": local_LayoutMode, "GutterPos": local_GutterPos, "BookFoldPrinting": local_BookFoldPrinting, "BookFoldRevPrinting": local_BookFoldRevPrinting, "BookFoldPrintingSheets": local_BookFoldPrintingSheets, }
	if (propertyName == "Windows"):
		retVal = this_Document.Windows
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_SyncScrollingSideBySide = retVal.SyncScrollingSideBySide
		except:
			local_SyncScrollingSideBySide = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Windows", "Count": local_Count, "SyncScrollingSideBySide": local_SyncScrollingSideBySide, }
	if (propertyName == "HasRoutingSlip"):
		retVal = this_Document.HasRoutingSlip
		return retVal
	if (propertyName == "RoutingSlip"):
		retVal = this_Document.RoutingSlip
		try:
			local_Subject = retVal.Subject
		except:
			local_Subject = None
		try:
			local_Message = retVal.Message
		except:
			local_Message = None
		try:
			local_Delivery = retVal.Delivery
		except:
			local_Delivery = None
		try:
			local_TrackStatus = retVal.TrackStatus
		except:
			local_TrackStatus = None
		try:
			local_Protect = retVal.Protect
		except:
			local_Protect = None
		try:
			local_ReturnWhenDone = retVal.ReturnWhenDone
		except:
			local_ReturnWhenDone = None
		try:
			local_Status = retVal.Status
		except:
			local_Status = None
		try:
			local_Recipients = retVal.Recipients
		except:
			local_Recipients = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "RoutingSlip", "Subject": local_Subject, "Message": local_Message, "Delivery": local_Delivery, "TrackStatus": local_TrackStatus, "Protect": local_Protect, "ReturnWhenDone": local_ReturnWhenDone, "Status": local_Status, "Recipients": local_Recipients, }
	if (propertyName == "Routed"):
		retVal = this_Document.Routed
		return retVal
	if (propertyName == "TablesOfAuthoritiesCategories"):
		retVal = this_Document.TablesOfAuthoritiesCategories
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "TablesOfAuthoritiesCategories", "Count": local_Count, }
	if (propertyName == "Indexes"):
		retVal = this_Document.Indexes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Format = retVal.Format
		except:
			local_Format = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Indexes", "Count": local_Count, "Format": local_Format, }
	if (propertyName == "Saved"):
		retVal = this_Document.Saved
		return retVal
	if (propertyName == "Content"):
		retVal = this_Document.Content
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_Start = retVal.Start
		except:
			local_Start = None
		try:
			local_End = retVal.End
		except:
			local_End = None
		try:
			local_StoryType = retVal.StoryType
		except:
			local_StoryType = None
		try:
			local_Bold = retVal.Bold
		except:
			local_Bold = None
		try:
			local_Italic = retVal.Italic
		except:
			local_Italic = None
		try:
			local_Underline = retVal.Underline
		except:
			local_Underline = None
		try:
			local_EmphasisMark = retVal.EmphasisMark
		except:
			local_EmphasisMark = None
		try:
			local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
		except:
			local_DisableCharacterSpaceGrid = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_StoryLength = retVal.StoryLength
		except:
			local_StoryLength = None
		try:
			local_LanguageID = retVal.LanguageID
		except:
			local_LanguageID = None
		try:
			local_GrammarChecked = retVal.GrammarChecked
		except:
			local_GrammarChecked = None
		try:
			local_SpellingChecked = retVal.SpellingChecked
		except:
			local_SpellingChecked = None
		try:
			local_HighlightColorIndex = retVal.HighlightColorIndex
		except:
			local_HighlightColorIndex = None
		try:
			local_CanEdit = retVal.CanEdit
		except:
			local_CanEdit = None
		try:
			local_CanPaste = retVal.CanPaste
		except:
			local_CanPaste = None
		try:
			local_IsEndOfRowMark = retVal.IsEndOfRowMark
		except:
			local_IsEndOfRowMark = None
		try:
			local_BookmarkID = retVal.BookmarkID
		except:
			local_BookmarkID = None
		try:
			local_PreviousBookmarkID = retVal.PreviousBookmarkID
		except:
			local_PreviousBookmarkID = None
		try:
			local_Case = retVal.Case
		except:
			local_Case = None
		try:
			local_Information = retVal.Information
		except:
			local_Information = None
		try:
			local_Orientation = retVal.Orientation
		except:
			local_Orientation = None
		try:
			local_LanguageIDFarEast = retVal.LanguageIDFarEast
		except:
			local_LanguageIDFarEast = None
		try:
			local_LanguageIDOther = retVal.LanguageIDOther
		except:
			local_LanguageIDOther = None
		try:
			local_LanguageDetected = retVal.LanguageDetected
		except:
			local_LanguageDetected = None
		try:
			local_FitTextWidth = retVal.FitTextWidth
		except:
			local_FitTextWidth = None
		try:
			local_HorizontalInVertical = retVal.HorizontalInVertical
		except:
			local_HorizontalInVertical = None
		try:
			local_TwoLinesInOne = retVal.TwoLinesInOne
		except:
			local_TwoLinesInOne = None
		try:
			local_CombineCharacters = retVal.CombineCharacters
		except:
			local_CombineCharacters = None
		try:
			local_NoProofing = retVal.NoProofing
		except:
			local_NoProofing = None
		try:
			local_CharacterWidth = retVal.CharacterWidth
		except:
			local_CharacterWidth = None
		try:
			local_Kana = retVal.Kana
		except:
			local_Kana = None
		try:
			local_BoldBi = retVal.BoldBi
		except:
			local_BoldBi = None
		try:
			local_ItalicBi = retVal.ItalicBi
		except:
			local_ItalicBi = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_ShowAll = retVal.ShowAll
		except:
			local_ShowAll = None
		try:
			local_CharacterStyle = retVal.CharacterStyle
		except:
			local_CharacterStyle = None
		try:
			local_ParagraphStyle = retVal.ParagraphStyle
		except:
			local_ParagraphStyle = None
		try:
			local_ListStyle = retVal.ListStyle
		except:
			local_ListStyle = None
		try:
			local_TableStyle = retVal.TableStyle
		except:
			local_TableStyle = None
		try:
			local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
		except:
			local_TextVisibleOnScreen = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }
	if (propertyName == "ActiveWindow"):
		retVal = this_Document.ActiveWindow
		try:
			local_Left = retVal.Left
		except:
			local_Left = None
		try:
			local_Top = retVal.Top
		except:
			local_Top = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_Split = retVal.Split
		except:
			local_Split = None
		try:
			local_SplitVertical = retVal.SplitVertical
		except:
			local_SplitVertical = None
		try:
			local_Caption = retVal.Caption
		except:
			local_Caption = None
		try:
			local_WindowState = retVal.WindowState
		except:
			local_WindowState = None
		try:
			local_DisplayRulers = retVal.DisplayRulers
		except:
			local_DisplayRulers = None
		try:
			local_DisplayVerticalRuler = retVal.DisplayVerticalRuler
		except:
			local_DisplayVerticalRuler = None
		try:
			local_Type = retVal.Type
		except:
			local_Type = None
		try:
			local_WindowNumber = retVal.WindowNumber
		except:
			local_WindowNumber = None
		try:
			local_DisplayVerticalScrollBar = retVal.DisplayVerticalScrollBar
		except:
			local_DisplayVerticalScrollBar = None
		try:
			local_DisplayHorizontalScrollBar = retVal.DisplayHorizontalScrollBar
		except:
			local_DisplayHorizontalScrollBar = None
		try:
			local_StyleAreaWidth = retVal.StyleAreaWidth
		except:
			local_StyleAreaWidth = None
		try:
			local_DisplayScreenTips = retVal.DisplayScreenTips
		except:
			local_DisplayScreenTips = None
		try:
			local_HorizontalPercentScrolled = retVal.HorizontalPercentScrolled
		except:
			local_HorizontalPercentScrolled = None
		try:
			local_VerticalPercentScrolled = retVal.VerticalPercentScrolled
		except:
			local_VerticalPercentScrolled = None
		try:
			local_DocumentMap = retVal.DocumentMap
		except:
			local_DocumentMap = None
		try:
			local_Active = retVal.Active
		except:
			local_Active = None
		try:
			local_DocumentMapPercentWidth = retVal.DocumentMapPercentWidth
		except:
			local_DocumentMapPercentWidth = None
		try:
			local_Index = retVal.Index
		except:
			local_Index = None
		try:
			local_IMEMode = retVal.IMEMode
		except:
			local_IMEMode = None
		try:
			local_UsableWidth = retVal.UsableWidth
		except:
			local_UsableWidth = None
		try:
			local_UsableHeight = retVal.UsableHeight
		except:
			local_UsableHeight = None
		try:
			local_EnvelopeVisible = retVal.EnvelopeVisible
		except:
			local_EnvelopeVisible = None
		try:
			local_DisplayRightRuler = retVal.DisplayRightRuler
		except:
			local_DisplayRightRuler = None
		try:
			local_DisplayLeftScrollBar = retVal.DisplayLeftScrollBar
		except:
			local_DisplayLeftScrollBar = None
		try:
			local_Visible = retVal.Visible
		except:
			local_Visible = None
		try:
			local_Thumbnails = retVal.Thumbnails
		except:
			local_Thumbnails = None
		try:
			local_ShowSourceDocuments = retVal.ShowSourceDocuments
		except:
			local_ShowSourceDocuments = None
		try:
			local_Hwnd = retVal.Hwnd
		except:
			local_Hwnd = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Window", "Left": local_Left, "Top": local_Top, "Width": local_Width, "Height": local_Height, "Split": local_Split, "SplitVertical": local_SplitVertical, "Caption": local_Caption, "WindowState": local_WindowState, "DisplayRulers": local_DisplayRulers, "DisplayVerticalRuler": local_DisplayVerticalRuler, "Type": local_Type, "WindowNumber": local_WindowNumber, "DisplayVerticalScrollBar": local_DisplayVerticalScrollBar, "DisplayHorizontalScrollBar": local_DisplayHorizontalScrollBar, "StyleAreaWidth": local_StyleAreaWidth, "DisplayScreenTips": local_DisplayScreenTips, "HorizontalPercentScrolled": local_HorizontalPercentScrolled, "VerticalPercentScrolled": local_VerticalPercentScrolled, "DocumentMap": local_DocumentMap, "Active": local_Active, "DocumentMapPercentWidth": local_DocumentMapPercentWidth, "Index": local_Index, "IMEMode": local_IMEMode, "UsableWidth": local_UsableWidth, "UsableHeight": local_UsableHeight, "EnvelopeVisible": local_EnvelopeVisible, "DisplayRightRuler": local_DisplayRightRuler, "DisplayLeftScrollBar": local_DisplayLeftScrollBar, "Visible": local_Visible, "Thumbnails": local_Thumbnails, "ShowSourceDocuments": local_ShowSourceDocuments, "Hwnd": local_Hwnd, }
	if (propertyName == "Kind"):
		retVal = this_Document.Kind
		return retVal
	if (propertyName == "ReadOnly"):
		retVal = this_Document.ReadOnly
		return retVal
	if (propertyName == "Subdocuments"):
		retVal = this_Document.Subdocuments
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Expanded = retVal.Expanded
		except:
			local_Expanded = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Subdocuments", "Count": local_Count, "Expanded": local_Expanded, }
	if (propertyName == "IsMasterDocument"):
		retVal = this_Document.IsMasterDocument
		return retVal
	if (propertyName == "DefaultTabStop"):
		retVal = this_Document.DefaultTabStop
		return retVal
	if (propertyName == "EmbedTrueTypeFonts"):
		retVal = this_Document.EmbedTrueTypeFonts
		return retVal
	if (propertyName == "SaveFormsData"):
		retVal = this_Document.SaveFormsData
		return retVal
	if (propertyName == "ReadOnlyRecommended"):
		retVal = this_Document.ReadOnlyRecommended
		return retVal
	if (propertyName == "SaveSubsetFonts"):
		retVal = this_Document.SaveSubsetFonts
		return retVal
	if (propertyName == "StoryRanges"):
		retVal = this_Document.StoryRanges
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "StoryRanges", "Count": local_Count, }
	if (propertyName == "CommandBars"):
		retVal = this_Document.CommandBars
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "CommandBars"}
	if (propertyName == "IsSubdocument"):
		retVal = this_Document.IsSubdocument
		return retVal
	if (propertyName == "SaveFormat"):
		retVal = this_Document.SaveFormat
		return retVal
	if (propertyName == "ProtectionType"):
		retVal = this_Document.ProtectionType
		return retVal
	if (propertyName == "Hyperlinks"):
		retVal = this_Document.Hyperlinks
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Hyperlinks", "Count": local_Count, }
	if (propertyName == "Shapes"):
		retVal = this_Document.Shapes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Shapes", "Count": local_Count, }
	if (propertyName == "ListTemplates"):
		retVal = this_Document.ListTemplates
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ListTemplates", "Count": local_Count, }
	if (propertyName == "Lists"):
		retVal = this_Document.Lists
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Lists", "Count": local_Count, }
	if (propertyName == "UpdateStylesOnOpen"):
		retVal = this_Document.UpdateStylesOnOpen
		return retVal
	if (propertyName == "AttachedTemplate"):
		retVal = this_Document.AttachedTemplate
		return retVal
	if (propertyName == "InlineShapes"):
		retVal = this_Document.InlineShapes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "InlineShapes", "Count": local_Count, }
	if (propertyName == "Background"):
		retVal = this_Document.Background
		try:
			local_AutoShapeType = retVal.AutoShapeType
		except:
			local_AutoShapeType = None
		try:
			local_ConnectionSiteCount = retVal.ConnectionSiteCount
		except:
			local_ConnectionSiteCount = None
		try:
			local_Connector = retVal.Connector
		except:
			local_Connector = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HorizontalFlip = retVal.HorizontalFlip
		except:
			local_HorizontalFlip = None
		try:
			local_Left = retVal.Left
		except:
			local_Left = None
		try:
			local_LockAspectRatio = retVal.LockAspectRatio
		except:
			local_LockAspectRatio = None
		try:
			local_Name = retVal.Name
		except:
			local_Name = None
		try:
			local_Rotation = retVal.Rotation
		except:
			local_Rotation = None
		try:
			local_Top = retVal.Top
		except:
			local_Top = None
		try:
			local_Type = retVal.Type
		except:
			local_Type = None
		try:
			local_VerticalFlip = retVal.VerticalFlip
		except:
			local_VerticalFlip = None
		try:
			local_Vertices = retVal.Vertices
		except:
			local_Vertices = None
		try:
			local_Visible = retVal.Visible
		except:
			local_Visible = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_ZOrderPosition = retVal.ZOrderPosition
		except:
			local_ZOrderPosition = None
		try:
			local_RelativeHorizontalPosition = retVal.RelativeHorizontalPosition
		except:
			local_RelativeHorizontalPosition = None
		try:
			local_RelativeVerticalPosition = retVal.RelativeVerticalPosition
		except:
			local_RelativeVerticalPosition = None
		try:
			local_LockAnchor = retVal.LockAnchor
		except:
			local_LockAnchor = None
		try:
			local_AlternativeText = retVal.AlternativeText
		except:
			local_AlternativeText = None
		try:
			local_HasDiagram = retVal.HasDiagram
		except:
			local_HasDiagram = None
		try:
			local_HasDiagramNode = retVal.HasDiagramNode
		except:
			local_HasDiagramNode = None
		try:
			local_Child = retVal.Child
		except:
			local_Child = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_LayoutInCell = retVal.LayoutInCell
		except:
			local_LayoutInCell = None
		try:
			local_HasChart = retVal.HasChart
		except:
			local_HasChart = None
		try:
			local_LeftRelative = retVal.LeftRelative
		except:
			local_LeftRelative = None
		try:
			local_TopRelative = retVal.TopRelative
		except:
			local_TopRelative = None
		try:
			local_WidthRelative = retVal.WidthRelative
		except:
			local_WidthRelative = None
		try:
			local_HeightRelative = retVal.HeightRelative
		except:
			local_HeightRelative = None
		try:
			local_RelativeHorizontalSize = retVal.RelativeHorizontalSize
		except:
			local_RelativeHorizontalSize = None
		try:
			local_RelativeVerticalSize = retVal.RelativeVerticalSize
		except:
			local_RelativeVerticalSize = None
		try:
			local_HasSmartArt = retVal.HasSmartArt
		except:
			local_HasSmartArt = None
		try:
			local_ShapeStyle = retVal.ShapeStyle
		except:
			local_ShapeStyle = None
		try:
			local_BackgroundStyle = retVal.BackgroundStyle
		except:
			local_BackgroundStyle = None
		try:
			local_Title = retVal.Title
		except:
			local_Title = None
		try:
			local_AnchorID = retVal.AnchorID
		except:
			local_AnchorID = None
		try:
			local_EditID = retVal.EditID
		except:
			local_EditID = None
		try:
			local_GraphicStyle = retVal.GraphicStyle
		except:
			local_GraphicStyle = None
		try:
			local_Decorative = retVal.Decorative
		except:
			local_Decorative = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Shape", "AutoShapeType": local_AutoShapeType, "ConnectionSiteCount": local_ConnectionSiteCount, "Connector": local_Connector, "Height": local_Height, "HorizontalFlip": local_HorizontalFlip, "Left": local_Left, "LockAspectRatio": local_LockAspectRatio, "Name": local_Name, "Rotation": local_Rotation, "Top": local_Top, "Type": local_Type, "VerticalFlip": local_VerticalFlip, "Vertices": local_Vertices, "Visible": local_Visible, "Width": local_Width, "ZOrderPosition": local_ZOrderPosition, "RelativeHorizontalPosition": local_RelativeHorizontalPosition, "RelativeVerticalPosition": local_RelativeVerticalPosition, "LockAnchor": local_LockAnchor, "AlternativeText": local_AlternativeText, "HasDiagram": local_HasDiagram, "HasDiagramNode": local_HasDiagramNode, "Child": local_Child, "ID": local_ID, "LayoutInCell": local_LayoutInCell, "HasChart": local_HasChart, "LeftRelative": local_LeftRelative, "TopRelative": local_TopRelative, "WidthRelative": local_WidthRelative, "HeightRelative": local_HeightRelative, "RelativeHorizontalSize": local_RelativeHorizontalSize, "RelativeVerticalSize": local_RelativeVerticalSize, "HasSmartArt": local_HasSmartArt, "ShapeStyle": local_ShapeStyle, "BackgroundStyle": local_BackgroundStyle, "Title": local_Title, "AnchorID": local_AnchorID, "EditID": local_EditID, "GraphicStyle": local_GraphicStyle, "Decorative": local_Decorative, }
	if (propertyName == "GrammarChecked"):
		retVal = this_Document.GrammarChecked
		return retVal
	if (propertyName == "SpellingChecked"):
		retVal = this_Document.SpellingChecked
		return retVal
	if (propertyName == "ShowGrammaticalErrors"):
		retVal = this_Document.ShowGrammaticalErrors
		return retVal
	if (propertyName == "ShowSpellingErrors"):
		retVal = this_Document.ShowSpellingErrors
		return retVal
	if (propertyName == "Versions"):
		retVal = this_Document.Versions
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_AutoVersion = retVal.AutoVersion
		except:
			local_AutoVersion = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Versions", "Count": local_Count, "AutoVersion": local_AutoVersion, }
	if (propertyName == "ShowSummary"):
		retVal = this_Document.ShowSummary
		return retVal
	if (propertyName == "SummaryViewMode"):
		retVal = this_Document.SummaryViewMode
		return retVal
	if (propertyName == "SummaryLength"):
		retVal = this_Document.SummaryLength
		return retVal
	if (propertyName == "PrintFractionalWidths"):
		retVal = this_Document.PrintFractionalWidths
		return retVal
	if (propertyName == "PrintPostScriptOverText"):
		retVal = this_Document.PrintPostScriptOverText
		return retVal
	if (propertyName == "Container"):
		retVal = this_Document.Container
		return retVal
	if (propertyName == "PrintFormsData"):
		retVal = this_Document.PrintFormsData
		return retVal
	if (propertyName == "ListParagraphs"):
		retVal = this_Document.ListParagraphs
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ListParagraphs", "Count": local_Count, }
	if (propertyName == "HasPassword"):
		retVal = this_Document.HasPassword
		return retVal
	if (propertyName == "WriteReserved"):
		retVal = this_Document.WriteReserved
		return retVal
	if (propertyName == "UserControl"):
		retVal = this_Document.UserControl
		return retVal
	if (propertyName == "HasMailer"):
		retVal = this_Document.HasMailer
		return retVal
	if (propertyName == "Mailer"):
		retVal = this_Document.Mailer
		try:
			local_BCCRecipients = retVal.BCCRecipients
		except:
			local_BCCRecipients = None
		try:
			local_CCRecipients = retVal.CCRecipients
		except:
			local_CCRecipients = None
		try:
			local_Recipients = retVal.Recipients
		except:
			local_Recipients = None
		try:
			local_Enclosures = retVal.Enclosures
		except:
			local_Enclosures = None
		try:
			local_Sender = retVal.Sender
		except:
			local_Sender = None
		try:
			local_SendDateTime = retVal.SendDateTime
		except:
			local_SendDateTime = None
		try:
			local_Received = retVal.Received
		except:
			local_Received = None
		try:
			local_Subject = retVal.Subject
		except:
			local_Subject = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Mailer", "BCCRecipients": local_BCCRecipients, "CCRecipients": local_CCRecipients, "Recipients": local_Recipients, "Enclosures": local_Enclosures, "Sender": local_Sender, "SendDateTime": local_SendDateTime, "Received": local_Received, "Subject": local_Subject, }
	if (propertyName == "ReadabilityStatistics"):
		retVal = this_Document.ReadabilityStatistics
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ReadabilityStatistics", "Count": local_Count, }
	if (propertyName == "GrammaticalErrors"):
		retVal = this_Document.GrammaticalErrors
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Type = retVal.Type
		except:
			local_Type = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ProofreadingErrors", "Count": local_Count, "Type": local_Type, }
	if (propertyName == "SpellingErrors"):
		retVal = this_Document.SpellingErrors
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Type = retVal.Type
		except:
			local_Type = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ProofreadingErrors", "Count": local_Count, "Type": local_Type, }
	if (propertyName == "VBProject"):
		retVal = this_Document.VBProject
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "VBProject"}
	if (propertyName == "FormsDesign"):
		retVal = this_Document.FormsDesign
		return retVal
	if (propertyName == "_CodeName"):
		retVal = this_Document._CodeName
		return retVal
	if (propertyName == "CodeName"):
		retVal = this_Document.CodeName
		return retVal
	if (propertyName == "SnapToGrid"):
		retVal = this_Document.SnapToGrid
		return retVal
	if (propertyName == "SnapToShapes"):
		retVal = this_Document.SnapToShapes
		return retVal
	if (propertyName == "GridDistanceHorizontal"):
		retVal = this_Document.GridDistanceHorizontal
		return retVal
	if (propertyName == "GridDistanceVertical"):
		retVal = this_Document.GridDistanceVertical
		return retVal
	if (propertyName == "GridOriginHorizontal"):
		retVal = this_Document.GridOriginHorizontal
		return retVal
	if (propertyName == "GridOriginVertical"):
		retVal = this_Document.GridOriginVertical
		return retVal
	if (propertyName == "GridSpaceBetweenHorizontalLines"):
		retVal = this_Document.GridSpaceBetweenHorizontalLines
		return retVal
	if (propertyName == "GridSpaceBetweenVerticalLines"):
		retVal = this_Document.GridSpaceBetweenVerticalLines
		return retVal
	if (propertyName == "GridOriginFromMargin"):
		retVal = this_Document.GridOriginFromMargin
		return retVal
	if (propertyName == "KerningByAlgorithm"):
		retVal = this_Document.KerningByAlgorithm
		return retVal
	if (propertyName == "JustificationMode"):
		retVal = this_Document.JustificationMode
		return retVal
	if (propertyName == "FarEastLineBreakLevel"):
		retVal = this_Document.FarEastLineBreakLevel
		return retVal
	if (propertyName == "NoLineBreakBefore"):
		retVal = this_Document.NoLineBreakBefore
		return retVal
	if (propertyName == "NoLineBreakAfter"):
		retVal = this_Document.NoLineBreakAfter
		return retVal
	if (propertyName == "TrackRevisions"):
		retVal = this_Document.TrackRevisions
		return retVal
	if (propertyName == "PrintRevisions"):
		retVal = this_Document.PrintRevisions
		return retVal
	if (propertyName == "ShowRevisions"):
		retVal = this_Document.ShowRevisions
		return retVal
	if (propertyName == "ActiveTheme"):
		retVal = this_Document.ActiveTheme
		return retVal
	if (propertyName == "ActiveThemeDisplayName"):
		retVal = this_Document.ActiveThemeDisplayName
		return retVal
	if (propertyName == "Email"):
		retVal = this_Document.Email
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Email"}
	if (propertyName == "Scripts"):
		retVal = this_Document.Scripts
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Scripts"}
	if (propertyName == "LanguageDetected"):
		retVal = this_Document.LanguageDetected
		return retVal
	if (propertyName == "FarEastLineBreakLanguage"):
		retVal = this_Document.FarEastLineBreakLanguage
		return retVal
	if (propertyName == "Frameset"):
		retVal = this_Document.Frameset
		try:
			local_Type = retVal.Type
		except:
			local_Type = None
		try:
			local_WidthType = retVal.WidthType
		except:
			local_WidthType = None
		try:
			local_HeightType = retVal.HeightType
		except:
			local_HeightType = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_ChildFramesetCount = retVal.ChildFramesetCount
		except:
			local_ChildFramesetCount = None
		try:
			local_FramesetBorderWidth = retVal.FramesetBorderWidth
		except:
			local_FramesetBorderWidth = None
		try:
			local_FramesetBorderColor = retVal.FramesetBorderColor
		except:
			local_FramesetBorderColor = None
		try:
			local_FrameScrollbarType = retVal.FrameScrollbarType
		except:
			local_FrameScrollbarType = None
		try:
			local_FrameResizable = retVal.FrameResizable
		except:
			local_FrameResizable = None
		try:
			local_FrameName = retVal.FrameName
		except:
			local_FrameName = None
		try:
			local_FrameDisplayBorders = retVal.FrameDisplayBorders
		except:
			local_FrameDisplayBorders = None
		try:
			local_FrameDefaultURL = retVal.FrameDefaultURL
		except:
			local_FrameDefaultURL = None
		try:
			local_FrameLinkToFile = retVal.FrameLinkToFile
		except:
			local_FrameLinkToFile = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Frameset", "Type": local_Type, "WidthType": local_WidthType, "HeightType": local_HeightType, "Width": local_Width, "Height": local_Height, "ChildFramesetCount": local_ChildFramesetCount, "FramesetBorderWidth": local_FramesetBorderWidth, "FramesetBorderColor": local_FramesetBorderColor, "FrameScrollbarType": local_FrameScrollbarType, "FrameResizable": local_FrameResizable, "FrameName": local_FrameName, "FrameDisplayBorders": local_FrameDisplayBorders, "FrameDefaultURL": local_FrameDefaultURL, "FrameLinkToFile": local_FrameLinkToFile, }
	if (propertyName == "ClickAndTypeParagraphStyle"):
		retVal = this_Document.ClickAndTypeParagraphStyle
		return retVal
	if (propertyName == "HTMLProject"):
		retVal = this_Document.HTMLProject
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "HTMLProject"}
	if (propertyName == "WebOptions"):
		retVal = this_Document.WebOptions
		try:
			local_OptimizeForBrowser = retVal.OptimizeForBrowser
		except:
			local_OptimizeForBrowser = None
		try:
			local_BrowserLevel = retVal.BrowserLevel
		except:
			local_BrowserLevel = None
		try:
			local_RelyOnCSS = retVal.RelyOnCSS
		except:
			local_RelyOnCSS = None
		try:
			local_OrganizeInFolder = retVal.OrganizeInFolder
		except:
			local_OrganizeInFolder = None
		try:
			local_UseLongFileNames = retVal.UseLongFileNames
		except:
			local_UseLongFileNames = None
		try:
			local_RelyOnVML = retVal.RelyOnVML
		except:
			local_RelyOnVML = None
		try:
			local_AllowPNG = retVal.AllowPNG
		except:
			local_AllowPNG = None
		try:
			local_ScreenSize = retVal.ScreenSize
		except:
			local_ScreenSize = None
		try:
			local_PixelsPerInch = retVal.PixelsPerInch
		except:
			local_PixelsPerInch = None
		try:
			local_Encoding = retVal.Encoding
		except:
			local_Encoding = None
		try:
			local_FolderSuffix = retVal.FolderSuffix
		except:
			local_FolderSuffix = None
		try:
			local_TargetBrowser = retVal.TargetBrowser
		except:
			local_TargetBrowser = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "WebOptions", "OptimizeForBrowser": local_OptimizeForBrowser, "BrowserLevel": local_BrowserLevel, "RelyOnCSS": local_RelyOnCSS, "OrganizeInFolder": local_OrganizeInFolder, "UseLongFileNames": local_UseLongFileNames, "RelyOnVML": local_RelyOnVML, "AllowPNG": local_AllowPNG, "ScreenSize": local_ScreenSize, "PixelsPerInch": local_PixelsPerInch, "Encoding": local_Encoding, "FolderSuffix": local_FolderSuffix, "TargetBrowser": local_TargetBrowser, }
	if (propertyName == "OpenEncoding"):
		retVal = this_Document.OpenEncoding
		return retVal
	if (propertyName == "SaveEncoding"):
		retVal = this_Document.SaveEncoding
		return retVal
	if (propertyName == "OptimizeForWord97"):
		retVal = this_Document.OptimizeForWord97
		return retVal
	if (propertyName == "VBASigned"):
		retVal = this_Document.VBASigned
		return retVal
	if (propertyName == "MailEnvelope"):
		retVal = this_Document.MailEnvelope
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "MsoEnvelope"}
	if (propertyName == "DisableFeatures"):
		retVal = this_Document.DisableFeatures
		return retVal
	if (propertyName == "DoNotEmbedSystemFonts"):
		retVal = this_Document.DoNotEmbedSystemFonts
		return retVal
	if (propertyName == "Signatures"):
		retVal = this_Document.Signatures
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "SignatureSet"}
	if (propertyName == "DefaultTargetFrame"):
		retVal = this_Document.DefaultTargetFrame
		return retVal
	if (propertyName == "HTMLDivisions"):
		retVal = this_Document.HTMLDivisions
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "HTMLDivisions", "Count": local_Count, "NestingLevel": local_NestingLevel, }
	if (propertyName == "DisableFeaturesIntroducedAfter"):
		retVal = this_Document.DisableFeaturesIntroducedAfter
		return retVal
	if (propertyName == "RemovePersonalInformation"):
		retVal = this_Document.RemovePersonalInformation
		return retVal
	if (propertyName == "SmartTags"):
		retVal = this_Document.SmartTags
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "SmartTags", "Count": local_Count, }
	if (propertyName == "EmbedSmartTags"):
		retVal = this_Document.EmbedSmartTags
		return retVal
	if (propertyName == "SmartTagsAsXMLProps"):
		retVal = this_Document.SmartTagsAsXMLProps
		return retVal
	if (propertyName == "TextEncoding"):
		retVal = this_Document.TextEncoding
		return retVal
	if (propertyName == "TextLineEnding"):
		retVal = this_Document.TextLineEnding
		return retVal
	if (propertyName == "StyleSheets"):
		retVal = this_Document.StyleSheets
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "StyleSheets", "Count": local_Count, }
	if (propertyName == "DefaultTableStyle"):
		retVal = this_Document.DefaultTableStyle
		return retVal
	if (propertyName == "PasswordEncryptionProvider"):
		retVal = this_Document.PasswordEncryptionProvider
		return retVal
	if (propertyName == "PasswordEncryptionAlgorithm"):
		retVal = this_Document.PasswordEncryptionAlgorithm
		return retVal
	if (propertyName == "PasswordEncryptionKeyLength"):
		retVal = this_Document.PasswordEncryptionKeyLength
		return retVal
	if (propertyName == "PasswordEncryptionFileProperties"):
		retVal = this_Document.PasswordEncryptionFileProperties
		return retVal
	if (propertyName == "EmbedLinguisticData"):
		retVal = this_Document.EmbedLinguisticData
		return retVal
	if (propertyName == "FormattingShowFont"):
		retVal = this_Document.FormattingShowFont
		return retVal
	if (propertyName == "FormattingShowClear"):
		retVal = this_Document.FormattingShowClear
		return retVal
	if (propertyName == "FormattingShowParagraph"):
		retVal = this_Document.FormattingShowParagraph
		return retVal
	if (propertyName == "FormattingShowNumbering"):
		retVal = this_Document.FormattingShowNumbering
		return retVal
	if (propertyName == "FormattingShowFilter"):
		retVal = this_Document.FormattingShowFilter
		return retVal
	if (propertyName == "Permission"):
		retVal = this_Document.Permission
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Permission"}
	if (propertyName == "XMLNodes"):
		retVal = this_Document.XMLNodes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "XMLNodes", "Count": local_Count, }
	if (propertyName == "XMLSchemaReferences"):
		retVal = this_Document.XMLSchemaReferences
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_AutomaticValidation = retVal.AutomaticValidation
		except:
			local_AutomaticValidation = None
		try:
			local_AllowSaveAsXMLWithoutValidation = retVal.AllowSaveAsXMLWithoutValidation
		except:
			local_AllowSaveAsXMLWithoutValidation = None
		try:
			local_HideValidationErrors = retVal.HideValidationErrors
		except:
			local_HideValidationErrors = None
		try:
			local_IgnoreMixedContent = retVal.IgnoreMixedContent
		except:
			local_IgnoreMixedContent = None
		try:
			local_ShowPlaceholderText = retVal.ShowPlaceholderText
		except:
			local_ShowPlaceholderText = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "XMLSchemaReferences", "Count": local_Count, "AutomaticValidation": local_AutomaticValidation, "AllowSaveAsXMLWithoutValidation": local_AllowSaveAsXMLWithoutValidation, "HideValidationErrors": local_HideValidationErrors, "IgnoreMixedContent": local_IgnoreMixedContent, "ShowPlaceholderText": local_ShowPlaceholderText, }
	if (propertyName == "SmartDocument"):
		retVal = this_Document.SmartDocument
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "SmartDocument"}
	if (propertyName == "SharedWorkspace"):
		retVal = this_Document.SharedWorkspace
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "SharedWorkspace"}
	if (propertyName == "Sync"):
		retVal = this_Document.Sync
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Sync"}
	if (propertyName == "EnforceStyle"):
		retVal = this_Document.EnforceStyle
		return retVal
	if (propertyName == "AutoFormatOverride"):
		retVal = this_Document.AutoFormatOverride
		return retVal
	if (propertyName == "XMLSaveDataOnly"):
		retVal = this_Document.XMLSaveDataOnly
		return retVal
	if (propertyName == "XMLHideNamespaces"):
		retVal = this_Document.XMLHideNamespaces
		return retVal
	if (propertyName == "XMLShowAdvancedErrors"):
		retVal = this_Document.XMLShowAdvancedErrors
		return retVal
	if (propertyName == "XMLUseXSLTWhenSaving"):
		retVal = this_Document.XMLUseXSLTWhenSaving
		return retVal
	if (propertyName == "XMLSaveThroughXSLT"):
		retVal = this_Document.XMLSaveThroughXSLT
		return retVal
	if (propertyName == "DocumentLibraryVersions"):
		retVal = this_Document.DocumentLibraryVersions
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "DocumentLibraryVersions"}
	if (propertyName == "ReadingModeLayoutFrozen"):
		retVal = this_Document.ReadingModeLayoutFrozen
		return retVal
	if (propertyName == "RemoveDateAndTime"):
		retVal = this_Document.RemoveDateAndTime
		return retVal
	if (propertyName == "ChildNodeSuggestions"):
		retVal = this_Document.ChildNodeSuggestions
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "XMLChildNodeSuggestions", "Count": local_Count, }
	if (propertyName == "XMLSchemaViolations"):
		retVal = this_Document.XMLSchemaViolations
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "XMLNodes", "Count": local_Count, }
	if (propertyName == "ReadingLayoutSizeX"):
		retVal = this_Document.ReadingLayoutSizeX
		return retVal
	if (propertyName == "ReadingLayoutSizeY"):
		retVal = this_Document.ReadingLayoutSizeY
		return retVal
	if (propertyName == "StyleSortMethod"):
		retVal = this_Document.StyleSortMethod
		return retVal
	if (propertyName == "ContentTypeProperties"):
		retVal = this_Document.ContentTypeProperties
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "MetaProperties"}
	if (propertyName == "TrackMoves"):
		retVal = this_Document.TrackMoves
		return retVal
	if (propertyName == "TrackFormatting"):
		retVal = this_Document.TrackFormatting
		return retVal
	if (propertyName == "OMaths"):
		retVal = this_Document.OMaths
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "OMaths", "Count": local_Count, }
	if (propertyName == "ServerPolicy"):
		retVal = this_Document.ServerPolicy
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ServerPolicy"}
	if (propertyName == "ContentControls"):
		retVal = this_Document.ContentControls
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ContentControls", "Count": local_Count, }
	if (propertyName == "DocumentInspectors"):
		retVal = this_Document.DocumentInspectors
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "DocumentInspectors"}
	if (propertyName == "Bibliography"):
		retVal = this_Document.Bibliography
		try:
			local_BibliographyStyle = retVal.BibliographyStyle
		except:
			local_BibliographyStyle = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Bibliography", "BibliographyStyle": local_BibliographyStyle, }
	if (propertyName == "LockTheme"):
		retVal = this_Document.LockTheme
		return retVal
	if (propertyName == "LockQuickStyleSet"):
		retVal = this_Document.LockQuickStyleSet
		return retVal
	if (propertyName == "OriginalDocumentTitle"):
		retVal = this_Document.OriginalDocumentTitle
		return retVal
	if (propertyName == "RevisedDocumentTitle"):
		retVal = this_Document.RevisedDocumentTitle
		return retVal
	if (propertyName == "CustomXMLParts"):
		retVal = this_Document.CustomXMLParts
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "CustomXMLParts"}
	if (propertyName == "FormattingShowNextLevel"):
		retVal = this_Document.FormattingShowNextLevel
		return retVal
	if (propertyName == "FormattingShowUserStyleName"):
		retVal = this_Document.FormattingShowUserStyleName
		return retVal
	if (propertyName == "Research"):
		retVal = this_Document.Research
		try:
			local_FavoriteService = retVal.FavoriteService
		except:
			local_FavoriteService = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Research", "FavoriteService": local_FavoriteService, }
	if (propertyName == "Final"):
		retVal = this_Document.Final
		return retVal
	if (propertyName == "OMathBreakBin"):
		retVal = this_Document.OMathBreakBin
		return retVal
	if (propertyName == "OMathBreakSub"):
		retVal = this_Document.OMathBreakSub
		return retVal
	if (propertyName == "OMathJc"):
		retVal = this_Document.OMathJc
		return retVal
	if (propertyName == "OMathLeftMargin"):
		retVal = this_Document.OMathLeftMargin
		return retVal
	if (propertyName == "OMathRightMargin"):
		retVal = this_Document.OMathRightMargin
		return retVal
	if (propertyName == "OMathWrap"):
		retVal = this_Document.OMathWrap
		return retVal
	if (propertyName == "OMathIntSubSupLim"):
		retVal = this_Document.OMathIntSubSupLim
		return retVal
	if (propertyName == "OMathNarySupSubLim"):
		retVal = this_Document.OMathNarySupSubLim
		return retVal
	if (propertyName == "OMathSmallFrac"):
		retVal = this_Document.OMathSmallFrac
		return retVal
	if (propertyName == "WordOpenXML"):
		retVal = this_Document.WordOpenXML
		return retVal
	if (propertyName == "DocumentTheme"):
		retVal = this_Document.DocumentTheme
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "OfficeTheme"}
	if (propertyName == "HasVBProject"):
		retVal = this_Document.HasVBProject
		return retVal
	if (propertyName == "OMathFontName"):
		retVal = this_Document.OMathFontName
		return retVal
	if (propertyName == "EncryptionProvider"):
		retVal = this_Document.EncryptionProvider
		return retVal
	if (propertyName == "UseMathDefaults"):
		retVal = this_Document.UseMathDefaults
		return retVal
	if (propertyName == "CurrentRsid"):
		retVal = this_Document.CurrentRsid
		return retVal
	if (propertyName == "DocID"):
		retVal = this_Document.DocID
		return retVal
	if (propertyName == "CompatibilityMode"):
		retVal = this_Document.CompatibilityMode
		return retVal
	if (propertyName == "CoAuthoring"):
		retVal = this_Document.CoAuthoring
		try:
			local_PendingUpdates = retVal.PendingUpdates
		except:
			local_PendingUpdates = None
		try:
			local_CanShare = retVal.CanShare
		except:
			local_CanShare = None
		try:
			local_CanMerge = retVal.CanMerge
		except:
			local_CanMerge = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "CoAuthoring", "PendingUpdates": local_PendingUpdates, "CanShare": local_CanShare, "CanMerge": local_CanMerge, }
	if (propertyName == "Broadcast"):
		retVal = this_Document.Broadcast
		try:
			local_AttendeeUrl = retVal.AttendeeUrl
		except:
			local_AttendeeUrl = None
		try:
			local_State = retVal.State
		except:
			local_State = None
		try:
			local_Capabilities = retVal.Capabilities
		except:
			local_Capabilities = None
		try:
			local_PresenterServiceUrl = retVal.PresenterServiceUrl
		except:
			local_PresenterServiceUrl = None
		try:
			local_SessionID = retVal.SessionID
		except:
			local_SessionID = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Broadcast", "AttendeeUrl": local_AttendeeUrl, "State": local_State, "Capabilities": local_Capabilities, "PresenterServiceUrl": local_PresenterServiceUrl, "SessionID": local_SessionID, }
	if (propertyName == "ChartDataPointTrack"):
		retVal = this_Document.ChartDataPointTrack
		return retVal
	if (propertyName == "IsInAutosave"):
		retVal = this_Document.IsInAutosave
		return retVal
	if (propertyName == "WorkIdentity"):
		retVal = this_Document.WorkIdentity
		return retVal
	if (propertyName == "AutoSaveOn"):
		retVal = this_Document.AutoSaveOn
		return retVal
	if (propertyName == "SensitivityLabel"):
		retVal = this_Document.SensitivityLabel
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ISensitivityLabel"}
	if (propertyName == "TrackJustMyRevisions"):
		retVal = this_Document.TrackJustMyRevisions
		return retVal


# Tool: 211
@mcp.tool()
async def word_Document_set_Property(this_Document_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Document
	
	propertyName: Name of the property. Can be one of ...
		AutoHyphenation, HyphenateCaps, HyphenationZone, ConsecutiveHyphensLimit, PageSetup, HasRoutingSlip, Saved, Kind, DefaultTabStop, EmbedTrueTypeFonts, SaveFormsData, ReadOnlyRecommended, SaveSubsetFonts, UpdateStylesOnOpen, AttachedTemplate, Background, GrammarChecked, SpellingChecked, ShowGrammaticalErrors, ShowSpellingErrors, ShowSummary, SummaryViewMode, SummaryLength, PrintFractionalWidths, PrintPostScriptOverText, PrintFormsData, Password, WritePassword, UserControl, HasMailer, _CodeName, SnapToGrid, SnapToShapes, GridDistanceHorizontal, GridDistanceVertical, GridOriginHorizontal, GridOriginVertical, GridSpaceBetweenHorizontalLines, GridSpaceBetweenVerticalLines, GridOriginFromMargin, KerningByAlgorithm, JustificationMode, FarEastLineBreakLevel, NoLineBreakBefore, NoLineBreakAfter, TrackRevisions, PrintRevisions, ShowRevisions, LanguageDetected, FarEastLineBreakLanguage, ClickAndTypeParagraphStyle, SaveEncoding, OptimizeForWord97, DisableFeatures, DoNotEmbedSystemFonts, DefaultTargetFrame, DisableFeaturesIntroducedAfter, RemovePersonalInformation, EmbedSmartTags, SmartTagsAsXMLProps, TextEncoding, TextLineEnding, EmbedLinguisticData, FormattingShowFont, FormattingShowClear, FormattingShowParagraph, FormattingShowNumbering, FormattingShowFilter, EnforceStyle, AutoFormatOverride, XMLSaveDataOnly, XMLHideNamespaces, XMLShowAdvancedErrors, XMLUseXSLTWhenSaving, XMLSaveThroughXSLT, ReadingModeLayoutFrozen, RemoveDateAndTime, ReadingLayoutSizeX, ReadingLayoutSizeY, StyleSortMethod, TrackMoves, TrackFormatting, LockTheme, LockQuickStyleSet, FormattingShowNextLevel, FormattingShowUserStyleName, Final, OMathBreakBin, OMathBreakSub, OMathJc, OMathLeftMargin, OMathRightMargin, OMathWrap, OMathIntSubSupLim, OMathNarySupSubLim, OMathSmallFrac, OMathFontName, EncryptionProvider, UseMathDefaults, ChartDataPointTrack, WorkIdentity, AutoSaveOn, TrackJustMyRevisions
	"""
	this_Document = get_object(this_Document_wordObjId)
	
	EnsureWord()
	if (propertyName == "AutoHyphenation"):
		this_Document.AutoHyphenation = propertyValue
	if (propertyName == "HyphenateCaps"):
		this_Document.HyphenateCaps = propertyValue
	if (propertyName == "HyphenationZone"):
		this_Document.HyphenationZone = propertyValue
	if (propertyName == "ConsecutiveHyphensLimit"):
		this_Document.ConsecutiveHyphensLimit = propertyValue
	if (propertyName == "PageSetup"):
		this_Document.PageSetup = propertyValue
	if (propertyName == "HasRoutingSlip"):
		this_Document.HasRoutingSlip = propertyValue
	if (propertyName == "Saved"):
		this_Document.Saved = propertyValue
	if (propertyName == "Kind"):
		this_Document.Kind = propertyValue
	if (propertyName == "DefaultTabStop"):
		this_Document.DefaultTabStop = propertyValue
	if (propertyName == "EmbedTrueTypeFonts"):
		this_Document.EmbedTrueTypeFonts = propertyValue
	if (propertyName == "SaveFormsData"):
		this_Document.SaveFormsData = propertyValue
	if (propertyName == "ReadOnlyRecommended"):
		this_Document.ReadOnlyRecommended = propertyValue
	if (propertyName == "SaveSubsetFonts"):
		this_Document.SaveSubsetFonts = propertyValue
	if (propertyName == "UpdateStylesOnOpen"):
		this_Document.UpdateStylesOnOpen = propertyValue
	if (propertyName == "AttachedTemplate"):
		this_Document.AttachedTemplate = propertyValue
	if (propertyName == "Background"):
		this_Document.Background = propertyValue
	if (propertyName == "GrammarChecked"):
		this_Document.GrammarChecked = propertyValue
	if (propertyName == "SpellingChecked"):
		this_Document.SpellingChecked = propertyValue
	if (propertyName == "ShowGrammaticalErrors"):
		this_Document.ShowGrammaticalErrors = propertyValue
	if (propertyName == "ShowSpellingErrors"):
		this_Document.ShowSpellingErrors = propertyValue
	if (propertyName == "ShowSummary"):
		this_Document.ShowSummary = propertyValue
	if (propertyName == "SummaryViewMode"):
		this_Document.SummaryViewMode = propertyValue
	if (propertyName == "SummaryLength"):
		this_Document.SummaryLength = propertyValue
	if (propertyName == "PrintFractionalWidths"):
		this_Document.PrintFractionalWidths = propertyValue
	if (propertyName == "PrintPostScriptOverText"):
		this_Document.PrintPostScriptOverText = propertyValue
	if (propertyName == "PrintFormsData"):
		this_Document.PrintFormsData = propertyValue
	if (propertyName == "Password"):
		this_Document.Password = propertyValue
	if (propertyName == "WritePassword"):
		this_Document.WritePassword = propertyValue
	if (propertyName == "UserControl"):
		this_Document.UserControl = propertyValue
	if (propertyName == "HasMailer"):
		this_Document.HasMailer = propertyValue
	if (propertyName == "_CodeName"):
		this_Document._CodeName = propertyValue
	if (propertyName == "SnapToGrid"):
		this_Document.SnapToGrid = propertyValue
	if (propertyName == "SnapToShapes"):
		this_Document.SnapToShapes = propertyValue
	if (propertyName == "GridDistanceHorizontal"):
		this_Document.GridDistanceHorizontal = propertyValue
	if (propertyName == "GridDistanceVertical"):
		this_Document.GridDistanceVertical = propertyValue
	if (propertyName == "GridOriginHorizontal"):
		this_Document.GridOriginHorizontal = propertyValue
	if (propertyName == "GridOriginVertical"):
		this_Document.GridOriginVertical = propertyValue
	if (propertyName == "GridSpaceBetweenHorizontalLines"):
		this_Document.GridSpaceBetweenHorizontalLines = propertyValue
	if (propertyName == "GridSpaceBetweenVerticalLines"):
		this_Document.GridSpaceBetweenVerticalLines = propertyValue
	if (propertyName == "GridOriginFromMargin"):
		this_Document.GridOriginFromMargin = propertyValue
	if (propertyName == "KerningByAlgorithm"):
		this_Document.KerningByAlgorithm = propertyValue
	if (propertyName == "JustificationMode"):
		this_Document.JustificationMode = propertyValue
	if (propertyName == "FarEastLineBreakLevel"):
		this_Document.FarEastLineBreakLevel = propertyValue
	if (propertyName == "NoLineBreakBefore"):
		this_Document.NoLineBreakBefore = propertyValue
	if (propertyName == "NoLineBreakAfter"):
		this_Document.NoLineBreakAfter = propertyValue
	if (propertyName == "TrackRevisions"):
		this_Document.TrackRevisions = propertyValue
	if (propertyName == "PrintRevisions"):
		this_Document.PrintRevisions = propertyValue
	if (propertyName == "ShowRevisions"):
		this_Document.ShowRevisions = propertyValue
	if (propertyName == "LanguageDetected"):
		this_Document.LanguageDetected = propertyValue
	if (propertyName == "FarEastLineBreakLanguage"):
		this_Document.FarEastLineBreakLanguage = propertyValue
	if (propertyName == "ClickAndTypeParagraphStyle"):
		this_Document.ClickAndTypeParagraphStyle = propertyValue
	if (propertyName == "SaveEncoding"):
		this_Document.SaveEncoding = propertyValue
	if (propertyName == "OptimizeForWord97"):
		this_Document.OptimizeForWord97 = propertyValue
	if (propertyName == "DisableFeatures"):
		this_Document.DisableFeatures = propertyValue
	if (propertyName == "DoNotEmbedSystemFonts"):
		this_Document.DoNotEmbedSystemFonts = propertyValue
	if (propertyName == "DefaultTargetFrame"):
		this_Document.DefaultTargetFrame = propertyValue
	if (propertyName == "DisableFeaturesIntroducedAfter"):
		this_Document.DisableFeaturesIntroducedAfter = propertyValue
	if (propertyName == "RemovePersonalInformation"):
		this_Document.RemovePersonalInformation = propertyValue
	if (propertyName == "EmbedSmartTags"):
		this_Document.EmbedSmartTags = propertyValue
	if (propertyName == "SmartTagsAsXMLProps"):
		this_Document.SmartTagsAsXMLProps = propertyValue
	if (propertyName == "TextEncoding"):
		this_Document.TextEncoding = propertyValue
	if (propertyName == "TextLineEnding"):
		this_Document.TextLineEnding = propertyValue
	if (propertyName == "EmbedLinguisticData"):
		this_Document.EmbedLinguisticData = propertyValue
	if (propertyName == "FormattingShowFont"):
		this_Document.FormattingShowFont = propertyValue
	if (propertyName == "FormattingShowClear"):
		this_Document.FormattingShowClear = propertyValue
	if (propertyName == "FormattingShowParagraph"):
		this_Document.FormattingShowParagraph = propertyValue
	if (propertyName == "FormattingShowNumbering"):
		this_Document.FormattingShowNumbering = propertyValue
	if (propertyName == "FormattingShowFilter"):
		this_Document.FormattingShowFilter = propertyValue
	if (propertyName == "EnforceStyle"):
		this_Document.EnforceStyle = propertyValue
	if (propertyName == "AutoFormatOverride"):
		this_Document.AutoFormatOverride = propertyValue
	if (propertyName == "XMLSaveDataOnly"):
		this_Document.XMLSaveDataOnly = propertyValue
	if (propertyName == "XMLHideNamespaces"):
		this_Document.XMLHideNamespaces = propertyValue
	if (propertyName == "XMLShowAdvancedErrors"):
		this_Document.XMLShowAdvancedErrors = propertyValue
	if (propertyName == "XMLUseXSLTWhenSaving"):
		this_Document.XMLUseXSLTWhenSaving = propertyValue
	if (propertyName == "XMLSaveThroughXSLT"):
		this_Document.XMLSaveThroughXSLT = propertyValue
	if (propertyName == "ReadingModeLayoutFrozen"):
		this_Document.ReadingModeLayoutFrozen = propertyValue
	if (propertyName == "RemoveDateAndTime"):
		this_Document.RemoveDateAndTime = propertyValue
	if (propertyName == "ReadingLayoutSizeX"):
		this_Document.ReadingLayoutSizeX = propertyValue
	if (propertyName == "ReadingLayoutSizeY"):
		this_Document.ReadingLayoutSizeY = propertyValue
	if (propertyName == "StyleSortMethod"):
		this_Document.StyleSortMethod = propertyValue
	if (propertyName == "TrackMoves"):
		this_Document.TrackMoves = propertyValue
	if (propertyName == "TrackFormatting"):
		this_Document.TrackFormatting = propertyValue
	if (propertyName == "LockTheme"):
		this_Document.LockTheme = propertyValue
	if (propertyName == "LockQuickStyleSet"):
		this_Document.LockQuickStyleSet = propertyValue
	if (propertyName == "FormattingShowNextLevel"):
		this_Document.FormattingShowNextLevel = propertyValue
	if (propertyName == "FormattingShowUserStyleName"):
		this_Document.FormattingShowUserStyleName = propertyValue
	if (propertyName == "Final"):
		this_Document.Final = propertyValue
	if (propertyName == "OMathBreakBin"):
		this_Document.OMathBreakBin = propertyValue
	if (propertyName == "OMathBreakSub"):
		this_Document.OMathBreakSub = propertyValue
	if (propertyName == "OMathJc"):
		this_Document.OMathJc = propertyValue
	if (propertyName == "OMathLeftMargin"):
		this_Document.OMathLeftMargin = propertyValue
	if (propertyName == "OMathRightMargin"):
		this_Document.OMathRightMargin = propertyValue
	if (propertyName == "OMathWrap"):
		this_Document.OMathWrap = propertyValue
	if (propertyName == "OMathIntSubSupLim"):
		this_Document.OMathIntSubSupLim = propertyValue
	if (propertyName == "OMathNarySupSubLim"):
		this_Document.OMathNarySupSubLim = propertyValue
	if (propertyName == "OMathSmallFrac"):
		this_Document.OMathSmallFrac = propertyValue
	if (propertyName == "OMathFontName"):
		this_Document.OMathFontName = propertyValue
	if (propertyName == "EncryptionProvider"):
		this_Document.EncryptionProvider = propertyValue
	if (propertyName == "UseMathDefaults"):
		this_Document.UseMathDefaults = propertyValue
	if (propertyName == "ChartDataPointTrack"):
		this_Document.ChartDataPointTrack = propertyValue
	if (propertyName == "WorkIdentity"):
		this_Document.WorkIdentity = propertyValue
	if (propertyName == "AutoSaveOn"):
		this_Document.AutoSaveOn = propertyValue
	if (propertyName == "TrackJustMyRevisions"):
		this_Document.TrackJustMyRevisions = propertyValue


# Tool: 212
@mcp.tool()
async def word_Range_Select(this_Range_wordObjId: str):
	"""This tool calls the Select methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.Select()


# Tool: 213
@mcp.tool()
async def word_Range_SetRange(this_Range_wordObjId: str, Start: int, End: int):
	"""This tool calls the SetRange methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Start: the Start as int
		End: the End as int
	"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.SetRange(Start, End)


# Tool: 214
@mcp.tool()
async def word_Range_Collapse(this_Range_wordObjId: str, Direction):
	"""This tool calls the Collapse methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Direction: the Direction as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Direction = tryParseString(Direction)
	this_Range.Collapse(Direction)


# Tool: 215
@mcp.tool()
async def word_Range_InsertBefore(this_Range_wordObjId: str, Text: str):
	"""This tool calls the InsertBefore methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Text: the Text as str
	"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.InsertBefore(Text)


# Tool: 216
@mcp.tool()
async def word_Range_InsertAfter(this_Range_wordObjId: str, Text: str):
	"""This tool calls the InsertAfter methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Text: the Text as str
	"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.InsertAfter(Text)


# Tool: 217
@mcp.tool()
async def word_Range_Next(this_Range_wordObjId: str, Unit, Count):
	"""This tool calls the Next methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	retVal = this_Range.Next(Unit, Count)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 218
@mcp.tool()
async def word_Range_Previous(this_Range_wordObjId: str, Unit, Count):
	"""This tool calls the Previous methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	retVal = this_Range.Previous(Unit, Count)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 219
@mcp.tool()
async def word_Range_StartOf(this_Range_wordObjId: str, Unit, Extend):
	"""This tool calls the StartOf methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Extend: the Extend as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Unit = tryParseString(Unit)
	Extend = tryParseString(Extend)
	retVal = this_Range.StartOf(Unit, Extend)
	return retVal


# Tool: 220
@mcp.tool()
async def word_Range_EndOf(this_Range_wordObjId: str, Unit, Extend):
	"""This tool calls the EndOf methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Extend: the Extend as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Unit = tryParseString(Unit)
	Extend = tryParseString(Extend)
	retVal = this_Range.EndOf(Unit, Extend)
	return retVal


# Tool: 221
@mcp.tool()
async def word_Range_Move(this_Range_wordObjId: str, Unit, Count):
	"""This tool calls the Move methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	retVal = this_Range.Move(Unit, Count)
	return retVal


# Tool: 222
@mcp.tool()
async def word_Range_MoveStart(this_Range_wordObjId: str, Unit, Count):
	"""This tool calls the MoveStart methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	retVal = this_Range.MoveStart(Unit, Count)
	return retVal


# Tool: 223
@mcp.tool()
async def word_Range_MoveEnd(this_Range_wordObjId: str, Unit, Count):
	"""This tool calls the MoveEnd methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	retVal = this_Range.MoveEnd(Unit, Count)
	return retVal


# Tool: 224
@mcp.tool()
async def word_Range_MoveWhile(this_Range_wordObjId: str, Cset, Count):
	"""This tool calls the MoveWhile methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Cset: the Cset as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Cset = tryParseString(Cset)
	Count = tryParseString(Count)
	retVal = this_Range.MoveWhile(Cset, Count)
	return retVal


# Tool: 225
@mcp.tool()
async def word_Range_MoveStartWhile(this_Range_wordObjId: str, Cset, Count):
	"""This tool calls the MoveStartWhile methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Cset: the Cset as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Cset = tryParseString(Cset)
	Count = tryParseString(Count)
	retVal = this_Range.MoveStartWhile(Cset, Count)
	return retVal


# Tool: 226
@mcp.tool()
async def word_Range_MoveEndWhile(this_Range_wordObjId: str, Cset, Count):
	"""This tool calls the MoveEndWhile methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Cset: the Cset as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Cset = tryParseString(Cset)
	Count = tryParseString(Count)
	retVal = this_Range.MoveEndWhile(Cset, Count)
	return retVal


# Tool: 227
@mcp.tool()
async def word_Range_MoveUntil(this_Range_wordObjId: str, Cset, Count):
	"""This tool calls the MoveUntil methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Cset: the Cset as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Cset = tryParseString(Cset)
	Count = tryParseString(Count)
	retVal = this_Range.MoveUntil(Cset, Count)
	return retVal


# Tool: 228
@mcp.tool()
async def word_Range_MoveStartUntil(this_Range_wordObjId: str, Cset, Count):
	"""This tool calls the MoveStartUntil methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Cset: the Cset as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Cset = tryParseString(Cset)
	Count = tryParseString(Count)
	retVal = this_Range.MoveStartUntil(Cset, Count)
	return retVal


# Tool: 229
@mcp.tool()
async def word_Range_MoveEndUntil(this_Range_wordObjId: str, Cset, Count):
	"""This tool calls the MoveEndUntil methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Cset: the Cset as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Cset = tryParseString(Cset)
	Count = tryParseString(Count)
	retVal = this_Range.MoveEndUntil(Cset, Count)
	return retVal


# Tool: 230
@mcp.tool()
async def word_Range_Cut(this_Range_wordObjId: str):
	"""This tool calls the Cut methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.Cut()


# Tool: 231
@mcp.tool()
async def word_Range_Copy(this_Range_wordObjId: str):
	"""This tool calls the Copy methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.Copy()


# Tool: 232
@mcp.tool()
async def word_Range_Paste(this_Range_wordObjId: str):
	"""This tool calls the Paste methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.Paste()


# Tool: 233
@mcp.tool()
async def word_Range_InsertBreak(this_Range_wordObjId: str, Type):
	"""This tool calls the InsertBreak methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Type: the Type as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Type = tryParseString(Type)
	this_Range.InsertBreak(Type)


# Tool: 234
@mcp.tool()
async def word_Range_InsertFile(this_Range_wordObjId: str, FileName: str, Range, ConfirmConversions, Link, Attachment):
	"""This tool calls the InsertFile methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as str
		Range: the Range as VT_VARIANT
		ConfirmConversions: the ConfirmConversions as VT_VARIANT
		Link: the Link as VT_VARIANT
		Attachment: the Attachment as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Range = tryParseString(Range)
	ConfirmConversions = tryParseString(ConfirmConversions)
	Link = tryParseString(Link)
	Attachment = tryParseString(Attachment)
	this_Range.InsertFile(FileName, Range, ConfirmConversions, Link, Attachment)


# Tool: 235
@mcp.tool()
async def word_Range_InStory(this_Range_wordObjId: str, Range_wordObjId: str):
	"""This tool calls the InStory methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Range_wordObjId: 		To pass this object, send in the __WordObjectId of the Range object as was obtained from a previous return value
	"""
	this_Range = get_object(this_Range_wordObjId)
	Range = get_object(Range_wordObjId)
	retVal = this_Range.InStory(Range)
	return retVal


# Tool: 236
@mcp.tool()
async def word_Range_InRange(this_Range_wordObjId: str, Range_wordObjId: str):
	"""This tool calls the InRange methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Range_wordObjId: 		To pass this object, send in the __WordObjectId of the Range object as was obtained from a previous return value
	"""
	this_Range = get_object(this_Range_wordObjId)
	Range = get_object(Range_wordObjId)
	retVal = this_Range.InRange(Range)
	return retVal


# Tool: 237
@mcp.tool()
async def word_Range_Delete(this_Range_wordObjId: str, Unit, Count):
	"""This tool calls the Delete methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	retVal = this_Range.Delete(Unit, Count)
	return retVal


# Tool: 238
@mcp.tool()
async def word_Range_WholeStory(this_Range_wordObjId: str):
	"""This tool calls the WholeStory methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.WholeStory()


# Tool: 239
@mcp.tool()
async def word_Range_Expand(this_Range_wordObjId: str, Unit):
	"""This tool calls the Expand methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Unit = tryParseString(Unit)
	retVal = this_Range.Expand(Unit)
	return retVal


# Tool: 240
@mcp.tool()
async def word_Range_InsertParagraph(this_Range_wordObjId: str):
	"""This tool calls the InsertParagraph methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.InsertParagraph()


# Tool: 241
@mcp.tool()
async def word_Range_InsertParagraphAfter(this_Range_wordObjId: str):
	"""This tool calls the InsertParagraphAfter methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.InsertParagraphAfter()


# Tool: 242
@mcp.tool()
async def word_Range_ConvertToTableOld(this_Range_wordObjId: str, Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit):
	"""This tool calls the ConvertToTableOld methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Separator: the Separator as VT_VARIANT
		NumRows: the NumRows as VT_VARIANT
		NumColumns: the NumColumns as VT_VARIANT
		InitialColumnWidth: the InitialColumnWidth as VT_VARIANT
		Format: the Format as VT_VARIANT
		ApplyBorders: the ApplyBorders as VT_VARIANT
		ApplyShading: the ApplyShading as VT_VARIANT
		ApplyFont: the ApplyFont as VT_VARIANT
		ApplyColor: the ApplyColor as VT_VARIANT
		ApplyHeadingRows: the ApplyHeadingRows as VT_VARIANT
		ApplyLastRow: the ApplyLastRow as VT_VARIANT
		ApplyFirstColumn: the ApplyFirstColumn as VT_VARIANT
		ApplyLastColumn: the ApplyLastColumn as VT_VARIANT
		AutoFit: the AutoFit as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Separator = tryParseString(Separator)
	NumRows = tryParseString(NumRows)
	NumColumns = tryParseString(NumColumns)
	InitialColumnWidth = tryParseString(InitialColumnWidth)
	Format = tryParseString(Format)
	ApplyBorders = tryParseString(ApplyBorders)
	ApplyShading = tryParseString(ApplyShading)
	ApplyFont = tryParseString(ApplyFont)
	ApplyColor = tryParseString(ApplyColor)
	ApplyHeadingRows = tryParseString(ApplyHeadingRows)
	ApplyLastRow = tryParseString(ApplyLastRow)
	ApplyFirstColumn = tryParseString(ApplyFirstColumn)
	ApplyLastColumn = tryParseString(ApplyLastColumn)
	AutoFit = tryParseString(AutoFit)
	retVal = this_Range.ConvertToTableOld(Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit)
	try:
		local_Uniform = retVal.Uniform
	except:
		local_Uniform = None
	try:
		local_AutoFormatType = retVal.AutoFormatType
	except:
		local_AutoFormatType = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_AllowPageBreaks = retVal.AllowPageBreaks
	except:
		local_AllowPageBreaks = None
	try:
		local_AllowAutoFit = retVal.AllowAutoFit
	except:
		local_AllowAutoFit = None
	try:
		local_PreferredWidth = retVal.PreferredWidth
	except:
		local_PreferredWidth = None
	try:
		local_PreferredWidthType = retVal.PreferredWidthType
	except:
		local_PreferredWidthType = None
	try:
		local_TopPadding = retVal.TopPadding
	except:
		local_TopPadding = None
	try:
		local_BottomPadding = retVal.BottomPadding
	except:
		local_BottomPadding = None
	try:
		local_LeftPadding = retVal.LeftPadding
	except:
		local_LeftPadding = None
	try:
		local_RightPadding = retVal.RightPadding
	except:
		local_RightPadding = None
	try:
		local_Spacing = retVal.Spacing
	except:
		local_Spacing = None
	try:
		local_TableDirection = retVal.TableDirection
	except:
		local_TableDirection = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_ApplyStyleHeadingRows = retVal.ApplyStyleHeadingRows
	except:
		local_ApplyStyleHeadingRows = None
	try:
		local_ApplyStyleLastRow = retVal.ApplyStyleLastRow
	except:
		local_ApplyStyleLastRow = None
	try:
		local_ApplyStyleFirstColumn = retVal.ApplyStyleFirstColumn
	except:
		local_ApplyStyleFirstColumn = None
	try:
		local_ApplyStyleLastColumn = retVal.ApplyStyleLastColumn
	except:
		local_ApplyStyleLastColumn = None
	try:
		local_ApplyStyleRowBands = retVal.ApplyStyleRowBands
	except:
		local_ApplyStyleRowBands = None
	try:
		local_ApplyStyleColumnBands = retVal.ApplyStyleColumnBands
	except:
		local_ApplyStyleColumnBands = None
	try:
		local_Title = retVal.Title
	except:
		local_Title = None
	try:
		local_Descr = retVal.Descr
	except:
		local_Descr = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Table", "Uniform": local_Uniform, "AutoFormatType": local_AutoFormatType, "NestingLevel": local_NestingLevel, "AllowPageBreaks": local_AllowPageBreaks, "AllowAutoFit": local_AllowAutoFit, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, "TopPadding": local_TopPadding, "BottomPadding": local_BottomPadding, "LeftPadding": local_LeftPadding, "RightPadding": local_RightPadding, "Spacing": local_Spacing, "TableDirection": local_TableDirection, "ID": local_ID, "Style": local_Style, "ApplyStyleHeadingRows": local_ApplyStyleHeadingRows, "ApplyStyleLastRow": local_ApplyStyleLastRow, "ApplyStyleFirstColumn": local_ApplyStyleFirstColumn, "ApplyStyleLastColumn": local_ApplyStyleLastColumn, "ApplyStyleRowBands": local_ApplyStyleRowBands, "ApplyStyleColumnBands": local_ApplyStyleColumnBands, "Title": local_Title, "Descr": local_Descr, }


# Tool: 243
@mcp.tool()
async def word_Range_InsertDateTimeOld(this_Range_wordObjId: str, DateTimeFormat, InsertAsField, InsertAsFullWidth):
	"""This tool calls the InsertDateTimeOld methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		DateTimeFormat: the DateTimeFormat as VT_VARIANT
		InsertAsField: the InsertAsField as VT_VARIANT
		InsertAsFullWidth: the InsertAsFullWidth as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	DateTimeFormat = tryParseString(DateTimeFormat)
	InsertAsField = tryParseString(InsertAsField)
	InsertAsFullWidth = tryParseString(InsertAsFullWidth)
	this_Range.InsertDateTimeOld(DateTimeFormat, InsertAsField, InsertAsFullWidth)


# Tool: 244
@mcp.tool()
async def word_Range_InsertSymbol(this_Range_wordObjId: str, CharacterNumber: int, Font, Unicode, Bias):
	"""This tool calls the InsertSymbol methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		CharacterNumber: the CharacterNumber as int
		Font: the Font as VT_VARIANT
		Unicode: the Unicode as VT_VARIANT
		Bias: the Bias as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Font = tryParseString(Font)
	Unicode = tryParseString(Unicode)
	Bias = tryParseString(Bias)
	this_Range.InsertSymbol(CharacterNumber, Font, Unicode, Bias)


# Tool: 245
@mcp.tool()
async def word_Range_InsertCrossReference_2002(this_Range_wordObjId: str, ReferenceType, ReferenceKind: int, ReferenceItem, InsertAsHyperlink, IncludePosition):
	"""This tool calls the InsertCrossReference_2002 methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		ReferenceType: the ReferenceType as VT_VARIANT
		ReferenceKind: the ReferenceKind as WdReferenceKind
		ReferenceItem: the ReferenceItem as VT_VARIANT
		InsertAsHyperlink: the InsertAsHyperlink as VT_VARIANT
		IncludePosition: the IncludePosition as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	ReferenceType = tryParseString(ReferenceType)
	ReferenceItem = tryParseString(ReferenceItem)
	InsertAsHyperlink = tryParseString(InsertAsHyperlink)
	IncludePosition = tryParseString(IncludePosition)
	this_Range.InsertCrossReference_2002(ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition)


# Tool: 246
@mcp.tool()
async def word_Range_InsertCaptionXP(this_Range_wordObjId: str, Label, Title, TitleAutoText, Position):
	"""This tool calls the InsertCaptionXP methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Label: the Label as VT_VARIANT
		Title: the Title as VT_VARIANT
		TitleAutoText: the TitleAutoText as VT_VARIANT
		Position: the Position as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Label = tryParseString(Label)
	Title = tryParseString(Title)
	TitleAutoText = tryParseString(TitleAutoText)
	Position = tryParseString(Position)
	this_Range.InsertCaptionXP(Label, Title, TitleAutoText, Position)


# Tool: 247
@mcp.tool()
async def word_Range_CopyAsPicture(this_Range_wordObjId: str):
	"""This tool calls the CopyAsPicture methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.CopyAsPicture()


# Tool: 248
@mcp.tool()
async def word_Range_SortOld(this_Range_wordObjId: str, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, LanguageID):
	"""This tool calls the SortOld methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		ExcludeHeader: the ExcludeHeader as VT_VARIANT
		FieldNumber: the FieldNumber as VT_VARIANT
		SortFieldType: the SortFieldType as VT_VARIANT
		SortOrder: the SortOrder as VT_VARIANT
		FieldNumber2: the FieldNumber2 as VT_VARIANT
		SortFieldType2: the SortFieldType2 as VT_VARIANT
		SortOrder2: the SortOrder2 as VT_VARIANT
		FieldNumber3: the FieldNumber3 as VT_VARIANT
		SortFieldType3: the SortFieldType3 as VT_VARIANT
		SortOrder3: the SortOrder3 as VT_VARIANT
		SortColumn: the SortColumn as VT_VARIANT
		Separator: the Separator as VT_VARIANT
		CaseSensitive: the CaseSensitive as VT_VARIANT
		LanguageID: the LanguageID as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	ExcludeHeader = tryParseString(ExcludeHeader)
	FieldNumber = tryParseString(FieldNumber)
	SortFieldType = tryParseString(SortFieldType)
	SortOrder = tryParseString(SortOrder)
	FieldNumber2 = tryParseString(FieldNumber2)
	SortFieldType2 = tryParseString(SortFieldType2)
	SortOrder2 = tryParseString(SortOrder2)
	FieldNumber3 = tryParseString(FieldNumber3)
	SortFieldType3 = tryParseString(SortFieldType3)
	SortOrder3 = tryParseString(SortOrder3)
	SortColumn = tryParseString(SortColumn)
	Separator = tryParseString(Separator)
	CaseSensitive = tryParseString(CaseSensitive)
	LanguageID = tryParseString(LanguageID)
	this_Range.SortOld(ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, LanguageID)


# Tool: 249
@mcp.tool()
async def word_Range_SortAscending(this_Range_wordObjId: str):
	"""This tool calls the SortAscending methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.SortAscending()


# Tool: 250
@mcp.tool()
async def word_Range_SortDescending(this_Range_wordObjId: str):
	"""This tool calls the SortDescending methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.SortDescending()


# Tool: 251
@mcp.tool()
async def word_Range_IsEqual(this_Range_wordObjId: str, Range_wordObjId: str):
	"""This tool calls the IsEqual methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Range_wordObjId: 		To pass this object, send in the __WordObjectId of the Range object as was obtained from a previous return value
	"""
	this_Range = get_object(this_Range_wordObjId)
	Range = get_object(Range_wordObjId)
	retVal = this_Range.IsEqual(Range)
	return retVal


# Tool: 252
@mcp.tool()
async def word_Range_Calculate(this_Range_wordObjId: str):
	"""This tool calls the Calculate methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	retVal = this_Range.Calculate()
	return retVal


# Tool: 253
@mcp.tool()
async def word_Range_GoTo(this_Range_wordObjId: str, What, Which, Count, Name):
	"""This tool calls the GoTo methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		What: the What as VT_VARIANT
		Which: the Which as VT_VARIANT
		Count: the Count as VT_VARIANT
		Name: the Name as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	What = tryParseString(What)
	Which = tryParseString(Which)
	Count = tryParseString(Count)
	Name = tryParseString(Name)
	retVal = this_Range.GoTo(What, Which, Count, Name)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 254
@mcp.tool()
async def word_Range_GoToNext(this_Range_wordObjId: str, What: int):
	"""This tool calls the GoToNext methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		What: the What as WdGoToItem
	"""
	this_Range = get_object(this_Range_wordObjId)
	retVal = this_Range.GoToNext(What)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 255
@mcp.tool()
async def word_Range_GoToPrevious(this_Range_wordObjId: str, What: int):
	"""This tool calls the GoToPrevious methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		What: the What as WdGoToItem
	"""
	this_Range = get_object(this_Range_wordObjId)
	retVal = this_Range.GoToPrevious(What)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 256
@mcp.tool()
async def word_Range_PasteSpecial(this_Range_wordObjId: str, IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel):
	"""This tool calls the PasteSpecial methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		IconIndex: the IconIndex as VT_VARIANT
		Link: the Link as VT_VARIANT
		Placement: the Placement as VT_VARIANT
		DisplayAsIcon: the DisplayAsIcon as VT_VARIANT
		DataType: the DataType as VT_VARIANT
		IconFileName: the IconFileName as VT_VARIANT
		IconLabel: the IconLabel as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	IconIndex = tryParseString(IconIndex)
	Link = tryParseString(Link)
	Placement = tryParseString(Placement)
	DisplayAsIcon = tryParseString(DisplayAsIcon)
	DataType = tryParseString(DataType)
	IconFileName = tryParseString(IconFileName)
	IconLabel = tryParseString(IconLabel)
	this_Range.PasteSpecial(IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel)


# Tool: 257
@mcp.tool()
async def word_Range_LookupNameProperties(this_Range_wordObjId: str):
	"""This tool calls the LookupNameProperties methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.LookupNameProperties()


# Tool: 258
@mcp.tool()
async def word_Range_ComputeStatistics(this_Range_wordObjId: str, Statistic: int):
	"""This tool calls the ComputeStatistics methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Statistic: the Statistic as WdStatistic
	"""
	this_Range = get_object(this_Range_wordObjId)
	retVal = this_Range.ComputeStatistics(Statistic)
	return retVal


# Tool: 259
@mcp.tool()
async def word_Range_Relocate(this_Range_wordObjId: str, Direction: int):
	"""This tool calls the Relocate methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Direction: the Direction as int
	"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.Relocate(Direction)


# Tool: 260
@mcp.tool()
async def word_Range_CheckSynonyms(this_Range_wordObjId: str):
	"""This tool calls the CheckSynonyms methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.CheckSynonyms()


# Tool: 261
@mcp.tool()
async def word_Range_SubscribeTo(this_Range_wordObjId: str, Edition: str, Format):
	"""This tool calls the SubscribeTo methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Edition: the Edition as str
		Format: the Format as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Format = tryParseString(Format)
	this_Range.SubscribeTo(Edition, Format)


# Tool: 262
@mcp.tool()
async def word_Range_CreatePublisher(this_Range_wordObjId: str, Edition, ContainsPICT, ContainsRTF, ContainsText):
	"""This tool calls the CreatePublisher methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Edition: the Edition as VT_VARIANT
		ContainsPICT: the ContainsPICT as VT_VARIANT
		ContainsRTF: the ContainsRTF as VT_VARIANT
		ContainsText: the ContainsText as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Edition = tryParseString(Edition)
	ContainsPICT = tryParseString(ContainsPICT)
	ContainsRTF = tryParseString(ContainsRTF)
	ContainsText = tryParseString(ContainsText)
	this_Range.CreatePublisher(Edition, ContainsPICT, ContainsRTF, ContainsText)


# Tool: 263
@mcp.tool()
async def word_Range_InsertAutoText(this_Range_wordObjId: str):
	"""This tool calls the InsertAutoText methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.InsertAutoText()


# Tool: 264
@mcp.tool()
async def word_Range_InsertDatabase(this_Range_wordObjId: str, Format, Style, LinkToSource, Connection, SQLStatement, SQLStatement1, PasswordDocument, PasswordTemplate, WritePasswordDocument, WritePasswordTemplate, DataSource, From, To, IncludeFields):
	"""This tool calls the InsertDatabase methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Format: the Format as VT_VARIANT
		Style: the Style as VT_VARIANT
		LinkToSource: the LinkToSource as VT_VARIANT
		Connection: the Connection as VT_VARIANT
		SQLStatement: the SQLStatement as VT_VARIANT
		SQLStatement1: the SQLStatement1 as VT_VARIANT
		PasswordDocument: the PasswordDocument as VT_VARIANT
		PasswordTemplate: the PasswordTemplate as VT_VARIANT
		WritePasswordDocument: the WritePasswordDocument as VT_VARIANT
		WritePasswordTemplate: the WritePasswordTemplate as VT_VARIANT
		DataSource: the DataSource as VT_VARIANT
		From: the From as VT_VARIANT
		To: the To as VT_VARIANT
		IncludeFields: the IncludeFields as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Format = tryParseString(Format)
	Style = tryParseString(Style)
	LinkToSource = tryParseString(LinkToSource)
	Connection = tryParseString(Connection)
	SQLStatement = tryParseString(SQLStatement)
	SQLStatement1 = tryParseString(SQLStatement1)
	PasswordDocument = tryParseString(PasswordDocument)
	PasswordTemplate = tryParseString(PasswordTemplate)
	WritePasswordDocument = tryParseString(WritePasswordDocument)
	WritePasswordTemplate = tryParseString(WritePasswordTemplate)
	DataSource = tryParseString(DataSource)
	From = tryParseString(From)
	To = tryParseString(To)
	IncludeFields = tryParseString(IncludeFields)
	this_Range.InsertDatabase(Format, Style, LinkToSource, Connection, SQLStatement, SQLStatement1, PasswordDocument, PasswordTemplate, WritePasswordDocument, WritePasswordTemplate, DataSource, From, To, IncludeFields)


# Tool: 265
@mcp.tool()
async def word_Range_AutoFormat(this_Range_wordObjId: str):
	"""This tool calls the AutoFormat methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.AutoFormat()


# Tool: 266
@mcp.tool()
async def word_Range_CheckGrammar(this_Range_wordObjId: str):
	"""This tool calls the CheckGrammar methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.CheckGrammar()


# Tool: 267
@mcp.tool()
async def word_Range_CheckSpelling(this_Range_wordObjId: str, CustomDictionary, IgnoreUppercase, AlwaysSuggest, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10):
	"""This tool calls the CheckSpelling methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		CustomDictionary: the CustomDictionary as VT_VARIANT
		IgnoreUppercase: the IgnoreUppercase as VT_VARIANT
		AlwaysSuggest: the AlwaysSuggest as VT_VARIANT
		CustomDictionary2: the CustomDictionary2 as VT_VARIANT
		CustomDictionary3: the CustomDictionary3 as VT_VARIANT
		CustomDictionary4: the CustomDictionary4 as VT_VARIANT
		CustomDictionary5: the CustomDictionary5 as VT_VARIANT
		CustomDictionary6: the CustomDictionary6 as VT_VARIANT
		CustomDictionary7: the CustomDictionary7 as VT_VARIANT
		CustomDictionary8: the CustomDictionary8 as VT_VARIANT
		CustomDictionary9: the CustomDictionary9 as VT_VARIANT
		CustomDictionary10: the CustomDictionary10 as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	CustomDictionary = tryParseString(CustomDictionary)
	IgnoreUppercase = tryParseString(IgnoreUppercase)
	AlwaysSuggest = tryParseString(AlwaysSuggest)
	CustomDictionary2 = tryParseString(CustomDictionary2)
	CustomDictionary3 = tryParseString(CustomDictionary3)
	CustomDictionary4 = tryParseString(CustomDictionary4)
	CustomDictionary5 = tryParseString(CustomDictionary5)
	CustomDictionary6 = tryParseString(CustomDictionary6)
	CustomDictionary7 = tryParseString(CustomDictionary7)
	CustomDictionary8 = tryParseString(CustomDictionary8)
	CustomDictionary9 = tryParseString(CustomDictionary9)
	CustomDictionary10 = tryParseString(CustomDictionary10)
	this_Range.CheckSpelling(CustomDictionary, IgnoreUppercase, AlwaysSuggest, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10)


# Tool: 268
@mcp.tool()
async def word_Range_GetSpellingSuggestions(this_Range_wordObjId: str, CustomDictionary, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10):
	"""This tool calls the GetSpellingSuggestions methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		CustomDictionary: the CustomDictionary as VT_VARIANT
		IgnoreUppercase: the IgnoreUppercase as VT_VARIANT
		MainDictionary: the MainDictionary as VT_VARIANT
		SuggestionMode: the SuggestionMode as VT_VARIANT
		CustomDictionary2: the CustomDictionary2 as VT_VARIANT
		CustomDictionary3: the CustomDictionary3 as VT_VARIANT
		CustomDictionary4: the CustomDictionary4 as VT_VARIANT
		CustomDictionary5: the CustomDictionary5 as VT_VARIANT
		CustomDictionary6: the CustomDictionary6 as VT_VARIANT
		CustomDictionary7: the CustomDictionary7 as VT_VARIANT
		CustomDictionary8: the CustomDictionary8 as VT_VARIANT
		CustomDictionary9: the CustomDictionary9 as VT_VARIANT
		CustomDictionary10: the CustomDictionary10 as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	CustomDictionary = tryParseString(CustomDictionary)
	IgnoreUppercase = tryParseString(IgnoreUppercase)
	MainDictionary = tryParseString(MainDictionary)
	SuggestionMode = tryParseString(SuggestionMode)
	CustomDictionary2 = tryParseString(CustomDictionary2)
	CustomDictionary3 = tryParseString(CustomDictionary3)
	CustomDictionary4 = tryParseString(CustomDictionary4)
	CustomDictionary5 = tryParseString(CustomDictionary5)
	CustomDictionary6 = tryParseString(CustomDictionary6)
	CustomDictionary7 = tryParseString(CustomDictionary7)
	CustomDictionary8 = tryParseString(CustomDictionary8)
	CustomDictionary9 = tryParseString(CustomDictionary9)
	CustomDictionary10 = tryParseString(CustomDictionary10)
	retVal = this_Range.GetSpellingSuggestions(CustomDictionary, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10)
	try:
		local_Count = retVal.Count
	except:
		local_Count = None
	try:
		local_SpellingErrorType = retVal.SpellingErrorType
	except:
		local_SpellingErrorType = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "SpellingSuggestions", "Count": local_Count, "SpellingErrorType": local_SpellingErrorType, }


# Tool: 269
@mcp.tool()
async def word_Range_InsertParagraphBefore(this_Range_wordObjId: str):
	"""This tool calls the InsertParagraphBefore methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.InsertParagraphBefore()


# Tool: 270
@mcp.tool()
async def word_Range_NextSubdocument(this_Range_wordObjId: str):
	"""This tool calls the NextSubdocument methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.NextSubdocument()


# Tool: 271
@mcp.tool()
async def word_Range_PreviousSubdocument(this_Range_wordObjId: str):
	"""This tool calls the PreviousSubdocument methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.PreviousSubdocument()


# Tool: 272
@mcp.tool()
async def word_Range_ConvertHangulAndHanja(this_Range_wordObjId: str, ConversionsMode, FastConversion, CheckHangulEnding, EnableRecentOrdering, CustomDictionary):
	"""This tool calls the ConvertHangulAndHanja methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		ConversionsMode: the ConversionsMode as VT_VARIANT
		FastConversion: the FastConversion as VT_VARIANT
		CheckHangulEnding: the CheckHangulEnding as VT_VARIANT
		EnableRecentOrdering: the EnableRecentOrdering as VT_VARIANT
		CustomDictionary: the CustomDictionary as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	ConversionsMode = tryParseString(ConversionsMode)
	FastConversion = tryParseString(FastConversion)
	CheckHangulEnding = tryParseString(CheckHangulEnding)
	EnableRecentOrdering = tryParseString(EnableRecentOrdering)
	CustomDictionary = tryParseString(CustomDictionary)
	this_Range.ConvertHangulAndHanja(ConversionsMode, FastConversion, CheckHangulEnding, EnableRecentOrdering, CustomDictionary)


# Tool: 273
@mcp.tool()
async def word_Range_PasteAsNestedTable(this_Range_wordObjId: str):
	"""This tool calls the PasteAsNestedTable methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.PasteAsNestedTable()


# Tool: 274
@mcp.tool()
async def word_Range_ModifyEnclosure(this_Range_wordObjId: str, Style, Symbol, EnclosedText):
	"""This tool calls the ModifyEnclosure methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Style: the Style as VT_VARIANT
		Symbol: the Symbol as VT_VARIANT
		EnclosedText: the EnclosedText as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Style = tryParseString(Style)
	Symbol = tryParseString(Symbol)
	EnclosedText = tryParseString(EnclosedText)
	this_Range.ModifyEnclosure(Style, Symbol, EnclosedText)


# Tool: 275
@mcp.tool()
async def word_Range_PhoneticGuide(this_Range_wordObjId: str, Text: str, Alignment: int, Raise: int, FontSize: int, FontName: str):
	"""This tool calls the PhoneticGuide methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Text: the Text as str
		Alignment: the Alignment as WdPhoneticGuideAlignmentType
		Raise: the Raise as int
		FontSize: the FontSize as int
		FontName: the FontName as str
	"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.PhoneticGuide(Text, Alignment, Raise, FontSize, FontName)


# Tool: 276
@mcp.tool()
async def word_Range_InsertDateTime(this_Range_wordObjId: str, DateTimeFormat, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType):
	"""This tool calls the InsertDateTime methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		DateTimeFormat: the DateTimeFormat as VT_VARIANT
		InsertAsField: the InsertAsField as VT_VARIANT
		InsertAsFullWidth: the InsertAsFullWidth as VT_VARIANT
		DateLanguage: the DateLanguage as VT_VARIANT
		CalendarType: the CalendarType as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	DateTimeFormat = tryParseString(DateTimeFormat)
	InsertAsField = tryParseString(InsertAsField)
	InsertAsFullWidth = tryParseString(InsertAsFullWidth)
	DateLanguage = tryParseString(DateLanguage)
	CalendarType = tryParseString(CalendarType)
	this_Range.InsertDateTime(DateTimeFormat, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType)


# Tool: 277
@mcp.tool()
async def word_Range_Sort(this_Range_wordObjId: str, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID):
	"""This tool calls the Sort methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		ExcludeHeader: the ExcludeHeader as VT_VARIANT
		FieldNumber: the FieldNumber as VT_VARIANT
		SortFieldType: the SortFieldType as VT_VARIANT
		SortOrder: the SortOrder as VT_VARIANT
		FieldNumber2: the FieldNumber2 as VT_VARIANT
		SortFieldType2: the SortFieldType2 as VT_VARIANT
		SortOrder2: the SortOrder2 as VT_VARIANT
		FieldNumber3: the FieldNumber3 as VT_VARIANT
		SortFieldType3: the SortFieldType3 as VT_VARIANT
		SortOrder3: the SortOrder3 as VT_VARIANT
		SortColumn: the SortColumn as VT_VARIANT
		Separator: the Separator as VT_VARIANT
		CaseSensitive: the CaseSensitive as VT_VARIANT
		BidiSort: the BidiSort as VT_VARIANT
		IgnoreThe: the IgnoreThe as VT_VARIANT
		IgnoreKashida: the IgnoreKashida as VT_VARIANT
		IgnoreDiacritics: the IgnoreDiacritics as VT_VARIANT
		IgnoreHe: the IgnoreHe as VT_VARIANT
		LanguageID: the LanguageID as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	ExcludeHeader = tryParseString(ExcludeHeader)
	FieldNumber = tryParseString(FieldNumber)
	SortFieldType = tryParseString(SortFieldType)
	SortOrder = tryParseString(SortOrder)
	FieldNumber2 = tryParseString(FieldNumber2)
	SortFieldType2 = tryParseString(SortFieldType2)
	SortOrder2 = tryParseString(SortOrder2)
	FieldNumber3 = tryParseString(FieldNumber3)
	SortFieldType3 = tryParseString(SortFieldType3)
	SortOrder3 = tryParseString(SortOrder3)
	SortColumn = tryParseString(SortColumn)
	Separator = tryParseString(Separator)
	CaseSensitive = tryParseString(CaseSensitive)
	BidiSort = tryParseString(BidiSort)
	IgnoreThe = tryParseString(IgnoreThe)
	IgnoreKashida = tryParseString(IgnoreKashida)
	IgnoreDiacritics = tryParseString(IgnoreDiacritics)
	IgnoreHe = tryParseString(IgnoreHe)
	LanguageID = tryParseString(LanguageID)
	this_Range.Sort(ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID)


# Tool: 278
@mcp.tool()
async def word_Range_DetectLanguage(this_Range_wordObjId: str):
	"""This tool calls the DetectLanguage methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.DetectLanguage()


# Tool: 279
@mcp.tool()
async def word_Range_ConvertToTable(this_Range_wordObjId: str, Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior):
	"""This tool calls the ConvertToTable methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Separator: the Separator as VT_VARIANT
		NumRows: the NumRows as VT_VARIANT
		NumColumns: the NumColumns as VT_VARIANT
		InitialColumnWidth: the InitialColumnWidth as VT_VARIANT
		Format: the Format as VT_VARIANT
		ApplyBorders: the ApplyBorders as VT_VARIANT
		ApplyShading: the ApplyShading as VT_VARIANT
		ApplyFont: the ApplyFont as VT_VARIANT
		ApplyColor: the ApplyColor as VT_VARIANT
		ApplyHeadingRows: the ApplyHeadingRows as VT_VARIANT
		ApplyLastRow: the ApplyLastRow as VT_VARIANT
		ApplyFirstColumn: the ApplyFirstColumn as VT_VARIANT
		ApplyLastColumn: the ApplyLastColumn as VT_VARIANT
		AutoFit: the AutoFit as VT_VARIANT
		AutoFitBehavior: the AutoFitBehavior as VT_VARIANT
		DefaultTableBehavior: the DefaultTableBehavior as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Separator = tryParseString(Separator)
	NumRows = tryParseString(NumRows)
	NumColumns = tryParseString(NumColumns)
	InitialColumnWidth = tryParseString(InitialColumnWidth)
	Format = tryParseString(Format)
	ApplyBorders = tryParseString(ApplyBorders)
	ApplyShading = tryParseString(ApplyShading)
	ApplyFont = tryParseString(ApplyFont)
	ApplyColor = tryParseString(ApplyColor)
	ApplyHeadingRows = tryParseString(ApplyHeadingRows)
	ApplyLastRow = tryParseString(ApplyLastRow)
	ApplyFirstColumn = tryParseString(ApplyFirstColumn)
	ApplyLastColumn = tryParseString(ApplyLastColumn)
	AutoFit = tryParseString(AutoFit)
	AutoFitBehavior = tryParseString(AutoFitBehavior)
	DefaultTableBehavior = tryParseString(DefaultTableBehavior)
	retVal = this_Range.ConvertToTable(Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior)
	try:
		local_Uniform = retVal.Uniform
	except:
		local_Uniform = None
	try:
		local_AutoFormatType = retVal.AutoFormatType
	except:
		local_AutoFormatType = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_AllowPageBreaks = retVal.AllowPageBreaks
	except:
		local_AllowPageBreaks = None
	try:
		local_AllowAutoFit = retVal.AllowAutoFit
	except:
		local_AllowAutoFit = None
	try:
		local_PreferredWidth = retVal.PreferredWidth
	except:
		local_PreferredWidth = None
	try:
		local_PreferredWidthType = retVal.PreferredWidthType
	except:
		local_PreferredWidthType = None
	try:
		local_TopPadding = retVal.TopPadding
	except:
		local_TopPadding = None
	try:
		local_BottomPadding = retVal.BottomPadding
	except:
		local_BottomPadding = None
	try:
		local_LeftPadding = retVal.LeftPadding
	except:
		local_LeftPadding = None
	try:
		local_RightPadding = retVal.RightPadding
	except:
		local_RightPadding = None
	try:
		local_Spacing = retVal.Spacing
	except:
		local_Spacing = None
	try:
		local_TableDirection = retVal.TableDirection
	except:
		local_TableDirection = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_ApplyStyleHeadingRows = retVal.ApplyStyleHeadingRows
	except:
		local_ApplyStyleHeadingRows = None
	try:
		local_ApplyStyleLastRow = retVal.ApplyStyleLastRow
	except:
		local_ApplyStyleLastRow = None
	try:
		local_ApplyStyleFirstColumn = retVal.ApplyStyleFirstColumn
	except:
		local_ApplyStyleFirstColumn = None
	try:
		local_ApplyStyleLastColumn = retVal.ApplyStyleLastColumn
	except:
		local_ApplyStyleLastColumn = None
	try:
		local_ApplyStyleRowBands = retVal.ApplyStyleRowBands
	except:
		local_ApplyStyleRowBands = None
	try:
		local_ApplyStyleColumnBands = retVal.ApplyStyleColumnBands
	except:
		local_ApplyStyleColumnBands = None
	try:
		local_Title = retVal.Title
	except:
		local_Title = None
	try:
		local_Descr = retVal.Descr
	except:
		local_Descr = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Table", "Uniform": local_Uniform, "AutoFormatType": local_AutoFormatType, "NestingLevel": local_NestingLevel, "AllowPageBreaks": local_AllowPageBreaks, "AllowAutoFit": local_AllowAutoFit, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, "TopPadding": local_TopPadding, "BottomPadding": local_BottomPadding, "LeftPadding": local_LeftPadding, "RightPadding": local_RightPadding, "Spacing": local_Spacing, "TableDirection": local_TableDirection, "ID": local_ID, "Style": local_Style, "ApplyStyleHeadingRows": local_ApplyStyleHeadingRows, "ApplyStyleLastRow": local_ApplyStyleLastRow, "ApplyStyleFirstColumn": local_ApplyStyleFirstColumn, "ApplyStyleLastColumn": local_ApplyStyleLastColumn, "ApplyStyleRowBands": local_ApplyStyleRowBands, "ApplyStyleColumnBands": local_ApplyStyleColumnBands, "Title": local_Title, "Descr": local_Descr, }


# Tool: 280
@mcp.tool()
async def word_Range_TCSCConverter(this_Range_wordObjId: str, WdTCSCConverterDirection: int, CommonTerms: bool, UseVariants: bool):
	"""This tool calls the TCSCConverter methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		WdTCSCConverterDirection: the WdTCSCConverterDirection as WdTCSCConverterDirection
		CommonTerms: the CommonTerms as bool
		UseVariants: the UseVariants as bool
	"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.TCSCConverter(WdTCSCConverterDirection, CommonTerms, UseVariants)


# Tool: 281
@mcp.tool()
async def word_Range_PasteAndFormat(this_Range_wordObjId: str, Type: int):
	"""This tool calls the PasteAndFormat methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Type: the Type as WdRecoveryType
	"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.PasteAndFormat(Type)


# Tool: 282
@mcp.tool()
async def word_Range_PasteExcelTable(this_Range_wordObjId: str, LinkedToExcel: bool, WordFormatting: bool, RTF: bool):
	"""This tool calls the PasteExcelTable methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		LinkedToExcel: the LinkedToExcel as bool
		WordFormatting: the WordFormatting as bool
		RTF: the RTF as bool
	"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.PasteExcelTable(LinkedToExcel, WordFormatting, RTF)


# Tool: 283
@mcp.tool()
async def word_Range_PasteAppendTable(this_Range_wordObjId: str):
	"""This tool calls the PasteAppendTable methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.PasteAppendTable()


# Tool: 284
@mcp.tool()
async def word_Range_GoToEditableRange(this_Range_wordObjId: str, EditorID):
	"""This tool calls the GoToEditableRange methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		EditorID: the EditorID as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	EditorID = tryParseString(EditorID)
	retVal = this_Range.GoToEditableRange(EditorID)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 285
@mcp.tool()
async def word_Range_InsertXML(this_Range_wordObjId: str, XML: str, Transform):
	"""This tool calls the InsertXML methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		XML: the XML as str
		Transform: the Transform as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Transform = tryParseString(Transform)
	this_Range.InsertXML(XML, Transform)


# Tool: 286
@mcp.tool()
async def word_Range_InsertCaption(this_Range_wordObjId: str, Label, Title, TitleAutoText, Position, ExcludeLabel):
	"""This tool calls the InsertCaption methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Label: the Label as VT_VARIANT
		Title: the Title as VT_VARIANT
		TitleAutoText: the TitleAutoText as VT_VARIANT
		Position: the Position as VT_VARIANT
		ExcludeLabel: the ExcludeLabel as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	Label = tryParseString(Label)
	Title = tryParseString(Title)
	TitleAutoText = tryParseString(TitleAutoText)
	Position = tryParseString(Position)
	ExcludeLabel = tryParseString(ExcludeLabel)
	this_Range.InsertCaption(Label, Title, TitleAutoText, Position, ExcludeLabel)


# Tool: 287
@mcp.tool()
async def word_Range_InsertCrossReference(this_Range_wordObjId: str, ReferenceType, ReferenceKind: int, ReferenceItem, InsertAsHyperlink, IncludePosition, SeparateNumbers, SeparatorString):
	"""This tool calls the InsertCrossReference methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		ReferenceType: the ReferenceType as VT_VARIANT
		ReferenceKind: the ReferenceKind as WdReferenceKind
		ReferenceItem: the ReferenceItem as VT_VARIANT
		InsertAsHyperlink: the InsertAsHyperlink as VT_VARIANT
		IncludePosition: the IncludePosition as VT_VARIANT
		SeparateNumbers: the SeparateNumbers as VT_VARIANT
		SeparatorString: the SeparatorString as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	ReferenceType = tryParseString(ReferenceType)
	ReferenceItem = tryParseString(ReferenceItem)
	InsertAsHyperlink = tryParseString(InsertAsHyperlink)
	IncludePosition = tryParseString(IncludePosition)
	SeparateNumbers = tryParseString(SeparateNumbers)
	SeparatorString = tryParseString(SeparatorString)
	this_Range.InsertCrossReference(ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition, SeparateNumbers, SeparatorString)


# Tool: 288
@mcp.tool()
async def word_Range_ExportFragment(this_Range_wordObjId: str, FileName: str, Format: int):
	"""This tool calls the ExportFragment methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as str
		Format: the Format as WdSaveFormat
	"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.ExportFragment(FileName, Format)


# Tool: 289
@mcp.tool()
async def word_Range_SetListLevel(this_Range_wordObjId: str, Level: int):
	"""This tool calls the SetListLevel methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Level: the Level as int
	"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.SetListLevel(Level)


# Tool: 290
@mcp.tool()
async def word_Range_InsertAlignmentTab(this_Range_wordObjId: str, Alignment: int, RelativeTo: int):
	"""This tool calls the InsertAlignmentTab methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		Alignment: the Alignment as int
		RelativeTo: the RelativeTo as int
	"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.InsertAlignmentTab(Alignment, RelativeTo)


# Tool: 291
@mcp.tool()
async def word_Range_ImportFragment(this_Range_wordObjId: str, FileName: str, MatchDestination: bool):
	"""This tool calls the ImportFragment methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as str
		MatchDestination: the MatchDestination as bool
	"""
	this_Range = get_object(this_Range_wordObjId)
	this_Range.ImportFragment(FileName, MatchDestination)


# Tool: 292
@mcp.tool()
async def word_Range_ExportAsFixedFormat(this_Range_wordObjId: str, OutputFileName: str, ExportFormat: int, OpenAfterExport: bool, OptimizeFor: int, ExportCurrentPage: bool, Item: int, IncludeDocProps: bool, KeepIRM: bool, CreateBookmarks: int, DocStructureTags: bool, BitmapMissingFonts: bool, UseISO19005_1: bool, FixedFormatExtClassPtr):
	"""This tool calls the ExportAsFixedFormat methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		OutputFileName: the OutputFileName as str
		ExportFormat: the ExportFormat as WdExportFormat
		OpenAfterExport: the OpenAfterExport as bool
		OptimizeFor: the OptimizeFor as WdExportOptimizeFor
		ExportCurrentPage: the ExportCurrentPage as bool
		Item: the Item as WdExportItem
		IncludeDocProps: the IncludeDocProps as bool
		KeepIRM: the KeepIRM as bool
		CreateBookmarks: the CreateBookmarks as WdExportCreateBookmarks
		DocStructureTags: the DocStructureTags as bool
		BitmapMissingFonts: the BitmapMissingFonts as bool
		UseISO19005_1: the UseISO19005_1 as bool
		FixedFormatExtClassPtr: the FixedFormatExtClassPtr as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	FixedFormatExtClassPtr = tryParseString(FixedFormatExtClassPtr)
	this_Range.ExportAsFixedFormat(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr)


# Tool: 293
@mcp.tool()
async def word_Range_SortByHeadings(this_Range_wordObjId: str, SortFieldType, SortOrder, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID):
	"""This tool calls the SortByHeadings methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		SortFieldType: the SortFieldType as VT_VARIANT
		SortOrder: the SortOrder as VT_VARIANT
		CaseSensitive: the CaseSensitive as VT_VARIANT
		BidiSort: the BidiSort as VT_VARIANT
		IgnoreThe: the IgnoreThe as VT_VARIANT
		IgnoreKashida: the IgnoreKashida as VT_VARIANT
		IgnoreDiacritics: the IgnoreDiacritics as VT_VARIANT
		IgnoreHe: the IgnoreHe as VT_VARIANT
		LanguageID: the LanguageID as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	SortFieldType = tryParseString(SortFieldType)
	SortOrder = tryParseString(SortOrder)
	CaseSensitive = tryParseString(CaseSensitive)
	BidiSort = tryParseString(BidiSort)
	IgnoreThe = tryParseString(IgnoreThe)
	IgnoreKashida = tryParseString(IgnoreKashida)
	IgnoreDiacritics = tryParseString(IgnoreDiacritics)
	IgnoreHe = tryParseString(IgnoreHe)
	LanguageID = tryParseString(LanguageID)
	this_Range.SortByHeadings(SortFieldType, SortOrder, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID)


# Tool: 294
@mcp.tool()
async def word_Range_ExportAsFixedFormat2(this_Range_wordObjId: str, OutputFileName: str, ExportFormat: int, OpenAfterExport: bool, OptimizeFor: int, ExportCurrentPage: bool, Item: int, IncludeDocProps: bool, KeepIRM: bool, CreateBookmarks: int, DocStructureTags: bool, BitmapMissingFonts: bool, UseISO19005_1: bool, OptimizeForImageQuality: bool, FixedFormatExtClassPtr):
	"""This tool calls the ExportAsFixedFormat2 methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		OutputFileName: the OutputFileName as str
		ExportFormat: the ExportFormat as WdExportFormat
		OpenAfterExport: the OpenAfterExport as bool
		OptimizeFor: the OptimizeFor as WdExportOptimizeFor
		ExportCurrentPage: the ExportCurrentPage as bool
		Item: the Item as WdExportItem
		IncludeDocProps: the IncludeDocProps as bool
		KeepIRM: the KeepIRM as bool
		CreateBookmarks: the CreateBookmarks as WdExportCreateBookmarks
		DocStructureTags: the DocStructureTags as bool
		BitmapMissingFonts: the BitmapMissingFonts as bool
		UseISO19005_1: the UseISO19005_1 as bool
		OptimizeForImageQuality: the OptimizeForImageQuality as bool
		FixedFormatExtClassPtr: the FixedFormatExtClassPtr as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	FixedFormatExtClassPtr = tryParseString(FixedFormatExtClassPtr)
	this_Range.ExportAsFixedFormat2(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, OptimizeForImageQuality, FixedFormatExtClassPtr)


# Tool: 295
@mcp.tool()
async def word_Range_ExportAsFixedFormat3(this_Range_wordObjId: str, OutputFileName: str, ExportFormat: int, OpenAfterExport: bool, OptimizeFor: int, ExportCurrentPage: bool, Item: int, IncludeDocProps: bool, KeepIRM: bool, CreateBookmarks: int, DocStructureTags: bool, BitmapMissingFonts: bool, UseISO19005_1: bool, OptimizeForImageQuality: bool, ImproveExportTagging: bool, FixedFormatExtClassPtr):
	"""This tool calls the ExportAsFixedFormat3 methodon an Range object. Pass the __WordObjectId of Range of the object you want to call the method on as the first parameter
	
	Parameters:
		OutputFileName: the OutputFileName as str
		ExportFormat: the ExportFormat as WdExportFormat
		OpenAfterExport: the OpenAfterExport as bool
		OptimizeFor: the OptimizeFor as WdExportOptimizeFor
		ExportCurrentPage: the ExportCurrentPage as bool
		Item: the Item as WdExportItem
		IncludeDocProps: the IncludeDocProps as bool
		KeepIRM: the KeepIRM as bool
		CreateBookmarks: the CreateBookmarks as WdExportCreateBookmarks
		DocStructureTags: the DocStructureTags as bool
		BitmapMissingFonts: the BitmapMissingFonts as bool
		UseISO19005_1: the UseISO19005_1 as bool
		OptimizeForImageQuality: the OptimizeForImageQuality as bool
		ImproveExportTagging: the ImproveExportTagging as bool
		FixedFormatExtClassPtr: the FixedFormatExtClassPtr as VT_VARIANT
	"""
	this_Range = get_object(this_Range_wordObjId)
	FixedFormatExtClassPtr = tryParseString(FixedFormatExtClassPtr)
	this_Range.ExportAsFixedFormat3(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, OptimizeForImageQuality, ImproveExportTagging, FixedFormatExtClassPtr)


# Tool: 296
@mcp.tool()
async def word_Range_get_Property(this_Range_wordObjId: str, propertyName: str):
	"""Gets properties of Range
	
	propertyName: Name of the property. Can be one of ...
		Text, FormattedText, Start, End, Font, Duplicate, StoryType, Tables, Words, Sentences, Characters, Footnotes, Endnotes, Comments, Cells, Sections, Paragraphs, Borders, Shading, TextRetrievalMode, Fields, FormFields, Frames, ParagraphFormat, ListFormat, Bookmarks, Bold, Italic, Underline, EmphasisMark, DisableCharacterSpaceGrid, Revisions, Style, StoryLength, LanguageID, SynonymInfo, Hyperlinks, ListParagraphs, Subdocuments, GrammarChecked, SpellingChecked, HighlightColorIndex, Columns, Rows, CanEdit, CanPaste, IsEndOfRowMark, BookmarkID, PreviousBookmarkID, Find, PageSetup, ShapeRange, Case, ReadabilityStatistics, GrammaticalErrors, SpellingErrors, Orientation, InlineShapes, NextStoryRange, LanguageIDFarEast, LanguageIDOther, LanguageDetected, FitTextWidth, HorizontalInVertical, TwoLinesInOne, CombineCharacters, NoProofing, TopLevelTables, Scripts, CharacterWidth, Kana, BoldBi, ItalicBi, ID, HTMLDivisions, SmartTags, ShowAll, Document, FootnoteOptions, EndnoteOptions, XMLNodes, XMLParentNode, Editors, EnhMetaFileBits, OMaths, CharacterStyle, ParagraphStyle, ListStyle, TableStyle, ContentControls, WordOpenXML, ParentContentControl, Locks, Updates, Conflicts, TextVisibleOnScreen
	"""
	this_Range = get_object(this_Range_wordObjId)
	
	EnsureWord()
	if (propertyName == "Text"):
		retVal = this_Range.Text
		return retVal
	if (propertyName == "FormattedText"):
		retVal = this_Range.FormattedText
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_Start = retVal.Start
		except:
			local_Start = None
		try:
			local_End = retVal.End
		except:
			local_End = None
		try:
			local_StoryType = retVal.StoryType
		except:
			local_StoryType = None
		try:
			local_Bold = retVal.Bold
		except:
			local_Bold = None
		try:
			local_Italic = retVal.Italic
		except:
			local_Italic = None
		try:
			local_Underline = retVal.Underline
		except:
			local_Underline = None
		try:
			local_EmphasisMark = retVal.EmphasisMark
		except:
			local_EmphasisMark = None
		try:
			local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
		except:
			local_DisableCharacterSpaceGrid = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_StoryLength = retVal.StoryLength
		except:
			local_StoryLength = None
		try:
			local_LanguageID = retVal.LanguageID
		except:
			local_LanguageID = None
		try:
			local_GrammarChecked = retVal.GrammarChecked
		except:
			local_GrammarChecked = None
		try:
			local_SpellingChecked = retVal.SpellingChecked
		except:
			local_SpellingChecked = None
		try:
			local_HighlightColorIndex = retVal.HighlightColorIndex
		except:
			local_HighlightColorIndex = None
		try:
			local_CanEdit = retVal.CanEdit
		except:
			local_CanEdit = None
		try:
			local_CanPaste = retVal.CanPaste
		except:
			local_CanPaste = None
		try:
			local_IsEndOfRowMark = retVal.IsEndOfRowMark
		except:
			local_IsEndOfRowMark = None
		try:
			local_BookmarkID = retVal.BookmarkID
		except:
			local_BookmarkID = None
		try:
			local_PreviousBookmarkID = retVal.PreviousBookmarkID
		except:
			local_PreviousBookmarkID = None
		try:
			local_Case = retVal.Case
		except:
			local_Case = None
		try:
			local_Information = retVal.Information
		except:
			local_Information = None
		try:
			local_Orientation = retVal.Orientation
		except:
			local_Orientation = None
		try:
			local_LanguageIDFarEast = retVal.LanguageIDFarEast
		except:
			local_LanguageIDFarEast = None
		try:
			local_LanguageIDOther = retVal.LanguageIDOther
		except:
			local_LanguageIDOther = None
		try:
			local_LanguageDetected = retVal.LanguageDetected
		except:
			local_LanguageDetected = None
		try:
			local_FitTextWidth = retVal.FitTextWidth
		except:
			local_FitTextWidth = None
		try:
			local_HorizontalInVertical = retVal.HorizontalInVertical
		except:
			local_HorizontalInVertical = None
		try:
			local_TwoLinesInOne = retVal.TwoLinesInOne
		except:
			local_TwoLinesInOne = None
		try:
			local_CombineCharacters = retVal.CombineCharacters
		except:
			local_CombineCharacters = None
		try:
			local_NoProofing = retVal.NoProofing
		except:
			local_NoProofing = None
		try:
			local_CharacterWidth = retVal.CharacterWidth
		except:
			local_CharacterWidth = None
		try:
			local_Kana = retVal.Kana
		except:
			local_Kana = None
		try:
			local_BoldBi = retVal.BoldBi
		except:
			local_BoldBi = None
		try:
			local_ItalicBi = retVal.ItalicBi
		except:
			local_ItalicBi = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_ShowAll = retVal.ShowAll
		except:
			local_ShowAll = None
		try:
			local_CharacterStyle = retVal.CharacterStyle
		except:
			local_CharacterStyle = None
		try:
			local_ParagraphStyle = retVal.ParagraphStyle
		except:
			local_ParagraphStyle = None
		try:
			local_ListStyle = retVal.ListStyle
		except:
			local_ListStyle = None
		try:
			local_TableStyle = retVal.TableStyle
		except:
			local_TableStyle = None
		try:
			local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
		except:
			local_TextVisibleOnScreen = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }
	if (propertyName == "Start"):
		retVal = this_Range.Start
		return retVal
	if (propertyName == "End"):
		retVal = this_Range.End
		return retVal
	if (propertyName == "Font"):
		retVal = this_Range.Font
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Font"}
	if (propertyName == "Duplicate"):
		retVal = this_Range.Duplicate
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_Start = retVal.Start
		except:
			local_Start = None
		try:
			local_End = retVal.End
		except:
			local_End = None
		try:
			local_StoryType = retVal.StoryType
		except:
			local_StoryType = None
		try:
			local_Bold = retVal.Bold
		except:
			local_Bold = None
		try:
			local_Italic = retVal.Italic
		except:
			local_Italic = None
		try:
			local_Underline = retVal.Underline
		except:
			local_Underline = None
		try:
			local_EmphasisMark = retVal.EmphasisMark
		except:
			local_EmphasisMark = None
		try:
			local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
		except:
			local_DisableCharacterSpaceGrid = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_StoryLength = retVal.StoryLength
		except:
			local_StoryLength = None
		try:
			local_LanguageID = retVal.LanguageID
		except:
			local_LanguageID = None
		try:
			local_GrammarChecked = retVal.GrammarChecked
		except:
			local_GrammarChecked = None
		try:
			local_SpellingChecked = retVal.SpellingChecked
		except:
			local_SpellingChecked = None
		try:
			local_HighlightColorIndex = retVal.HighlightColorIndex
		except:
			local_HighlightColorIndex = None
		try:
			local_CanEdit = retVal.CanEdit
		except:
			local_CanEdit = None
		try:
			local_CanPaste = retVal.CanPaste
		except:
			local_CanPaste = None
		try:
			local_IsEndOfRowMark = retVal.IsEndOfRowMark
		except:
			local_IsEndOfRowMark = None
		try:
			local_BookmarkID = retVal.BookmarkID
		except:
			local_BookmarkID = None
		try:
			local_PreviousBookmarkID = retVal.PreviousBookmarkID
		except:
			local_PreviousBookmarkID = None
		try:
			local_Case = retVal.Case
		except:
			local_Case = None
		try:
			local_Information = retVal.Information
		except:
			local_Information = None
		try:
			local_Orientation = retVal.Orientation
		except:
			local_Orientation = None
		try:
			local_LanguageIDFarEast = retVal.LanguageIDFarEast
		except:
			local_LanguageIDFarEast = None
		try:
			local_LanguageIDOther = retVal.LanguageIDOther
		except:
			local_LanguageIDOther = None
		try:
			local_LanguageDetected = retVal.LanguageDetected
		except:
			local_LanguageDetected = None
		try:
			local_FitTextWidth = retVal.FitTextWidth
		except:
			local_FitTextWidth = None
		try:
			local_HorizontalInVertical = retVal.HorizontalInVertical
		except:
			local_HorizontalInVertical = None
		try:
			local_TwoLinesInOne = retVal.TwoLinesInOne
		except:
			local_TwoLinesInOne = None
		try:
			local_CombineCharacters = retVal.CombineCharacters
		except:
			local_CombineCharacters = None
		try:
			local_NoProofing = retVal.NoProofing
		except:
			local_NoProofing = None
		try:
			local_CharacterWidth = retVal.CharacterWidth
		except:
			local_CharacterWidth = None
		try:
			local_Kana = retVal.Kana
		except:
			local_Kana = None
		try:
			local_BoldBi = retVal.BoldBi
		except:
			local_BoldBi = None
		try:
			local_ItalicBi = retVal.ItalicBi
		except:
			local_ItalicBi = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_ShowAll = retVal.ShowAll
		except:
			local_ShowAll = None
		try:
			local_CharacterStyle = retVal.CharacterStyle
		except:
			local_CharacterStyle = None
		try:
			local_ParagraphStyle = retVal.ParagraphStyle
		except:
			local_ParagraphStyle = None
		try:
			local_ListStyle = retVal.ListStyle
		except:
			local_ListStyle = None
		try:
			local_TableStyle = retVal.TableStyle
		except:
			local_TableStyle = None
		try:
			local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
		except:
			local_TextVisibleOnScreen = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }
	if (propertyName == "StoryType"):
		retVal = this_Range.StoryType
		return retVal
	if (propertyName == "Tables"):
		retVal = this_Range.Tables
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Tables", "Count": local_Count, "NestingLevel": local_NestingLevel, }
	if (propertyName == "Words"):
		retVal = this_Range.Words
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Words", "Count": local_Count, }
	if (propertyName == "Sentences"):
		retVal = this_Range.Sentences
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Sentences", "Count": local_Count, }
	if (propertyName == "Characters"):
		retVal = this_Range.Characters
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Characters", "Count": local_Count, }
	if (propertyName == "Footnotes"):
		retVal = this_Range.Footnotes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Location = retVal.Location
		except:
			local_Location = None
		try:
			local_NumberStyle = retVal.NumberStyle
		except:
			local_NumberStyle = None
		try:
			local_StartingNumber = retVal.StartingNumber
		except:
			local_StartingNumber = None
		try:
			local_NumberingRule = retVal.NumberingRule
		except:
			local_NumberingRule = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Footnotes", "Count": local_Count, "Location": local_Location, "NumberStyle": local_NumberStyle, "StartingNumber": local_StartingNumber, "NumberingRule": local_NumberingRule, }
	if (propertyName == "Endnotes"):
		retVal = this_Range.Endnotes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Location = retVal.Location
		except:
			local_Location = None
		try:
			local_NumberStyle = retVal.NumberStyle
		except:
			local_NumberStyle = None
		try:
			local_StartingNumber = retVal.StartingNumber
		except:
			local_StartingNumber = None
		try:
			local_NumberingRule = retVal.NumberingRule
		except:
			local_NumberingRule = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Endnotes", "Count": local_Count, "Location": local_Location, "NumberStyle": local_NumberStyle, "StartingNumber": local_StartingNumber, "NumberingRule": local_NumberingRule, }
	if (propertyName == "Comments"):
		retVal = this_Range.Comments
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_ShowBy = retVal.ShowBy
		except:
			local_ShowBy = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Comments", "Count": local_Count, "ShowBy": local_ShowBy, }
	if (propertyName == "Cells"):
		retVal = this_Range.Cells
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_VerticalAlignment = retVal.VerticalAlignment
		except:
			local_VerticalAlignment = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Cells", "Count": local_Count, "Width": local_Width, "Height": local_Height, "HeightRule": local_HeightRule, "VerticalAlignment": local_VerticalAlignment, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "Sections"):
		retVal = this_Range.Sections
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Sections", "Count": local_Count, }
	if (propertyName == "Paragraphs"):
		retVal = this_Range.Paragraphs
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_Alignment = retVal.Alignment
		except:
			local_Alignment = None
		try:
			local_KeepTogether = retVal.KeepTogether
		except:
			local_KeepTogether = None
		try:
			local_KeepWithNext = retVal.KeepWithNext
		except:
			local_KeepWithNext = None
		try:
			local_PageBreakBefore = retVal.PageBreakBefore
		except:
			local_PageBreakBefore = None
		try:
			local_NoLineNumber = retVal.NoLineNumber
		except:
			local_NoLineNumber = None
		try:
			local_RightIndent = retVal.RightIndent
		except:
			local_RightIndent = None
		try:
			local_LeftIndent = retVal.LeftIndent
		except:
			local_LeftIndent = None
		try:
			local_FirstLineIndent = retVal.FirstLineIndent
		except:
			local_FirstLineIndent = None
		try:
			local_LineSpacing = retVal.LineSpacing
		except:
			local_LineSpacing = None
		try:
			local_LineSpacingRule = retVal.LineSpacingRule
		except:
			local_LineSpacingRule = None
		try:
			local_SpaceBefore = retVal.SpaceBefore
		except:
			local_SpaceBefore = None
		try:
			local_SpaceAfter = retVal.SpaceAfter
		except:
			local_SpaceAfter = None
		try:
			local_Hyphenation = retVal.Hyphenation
		except:
			local_Hyphenation = None
		try:
			local_WidowControl = retVal.WidowControl
		except:
			local_WidowControl = None
		try:
			local_FarEastLineBreakControl = retVal.FarEastLineBreakControl
		except:
			local_FarEastLineBreakControl = None
		try:
			local_WordWrap = retVal.WordWrap
		except:
			local_WordWrap = None
		try:
			local_HangingPunctuation = retVal.HangingPunctuation
		except:
			local_HangingPunctuation = None
		try:
			local_HalfWidthPunctuationOnTopOfLine = retVal.HalfWidthPunctuationOnTopOfLine
		except:
			local_HalfWidthPunctuationOnTopOfLine = None
		try:
			local_AddSpaceBetweenFarEastAndAlpha = retVal.AddSpaceBetweenFarEastAndAlpha
		except:
			local_AddSpaceBetweenFarEastAndAlpha = None
		try:
			local_AddSpaceBetweenFarEastAndDigit = retVal.AddSpaceBetweenFarEastAndDigit
		except:
			local_AddSpaceBetweenFarEastAndDigit = None
		try:
			local_BaseLineAlignment = retVal.BaseLineAlignment
		except:
			local_BaseLineAlignment = None
		try:
			local_AutoAdjustRightIndent = retVal.AutoAdjustRightIndent
		except:
			local_AutoAdjustRightIndent = None
		try:
			local_DisableLineHeightGrid = retVal.DisableLineHeightGrid
		except:
			local_DisableLineHeightGrid = None
		try:
			local_OutlineLevel = retVal.OutlineLevel
		except:
			local_OutlineLevel = None
		try:
			local_CharacterUnitRightIndent = retVal.CharacterUnitRightIndent
		except:
			local_CharacterUnitRightIndent = None
		try:
			local_CharacterUnitLeftIndent = retVal.CharacterUnitLeftIndent
		except:
			local_CharacterUnitLeftIndent = None
		try:
			local_CharacterUnitFirstLineIndent = retVal.CharacterUnitFirstLineIndent
		except:
			local_CharacterUnitFirstLineIndent = None
		try:
			local_LineUnitBefore = retVal.LineUnitBefore
		except:
			local_LineUnitBefore = None
		try:
			local_LineUnitAfter = retVal.LineUnitAfter
		except:
			local_LineUnitAfter = None
		try:
			local_ReadingOrder = retVal.ReadingOrder
		except:
			local_ReadingOrder = None
		try:
			local_SpaceBeforeAuto = retVal.SpaceBeforeAuto
		except:
			local_SpaceBeforeAuto = None
		try:
			local_SpaceAfterAuto = retVal.SpaceAfterAuto
		except:
			local_SpaceAfterAuto = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Paragraphs", "Count": local_Count, "Style": local_Style, "Alignment": local_Alignment, "KeepTogether": local_KeepTogether, "KeepWithNext": local_KeepWithNext, "PageBreakBefore": local_PageBreakBefore, "NoLineNumber": local_NoLineNumber, "RightIndent": local_RightIndent, "LeftIndent": local_LeftIndent, "FirstLineIndent": local_FirstLineIndent, "LineSpacing": local_LineSpacing, "LineSpacingRule": local_LineSpacingRule, "SpaceBefore": local_SpaceBefore, "SpaceAfter": local_SpaceAfter, "Hyphenation": local_Hyphenation, "WidowControl": local_WidowControl, "FarEastLineBreakControl": local_FarEastLineBreakControl, "WordWrap": local_WordWrap, "HangingPunctuation": local_HangingPunctuation, "HalfWidthPunctuationOnTopOfLine": local_HalfWidthPunctuationOnTopOfLine, "AddSpaceBetweenFarEastAndAlpha": local_AddSpaceBetweenFarEastAndAlpha, "AddSpaceBetweenFarEastAndDigit": local_AddSpaceBetweenFarEastAndDigit, "BaseLineAlignment": local_BaseLineAlignment, "AutoAdjustRightIndent": local_AutoAdjustRightIndent, "DisableLineHeightGrid": local_DisableLineHeightGrid, "OutlineLevel": local_OutlineLevel, "CharacterUnitRightIndent": local_CharacterUnitRightIndent, "CharacterUnitLeftIndent": local_CharacterUnitLeftIndent, "CharacterUnitFirstLineIndent": local_CharacterUnitFirstLineIndent, "LineUnitBefore": local_LineUnitBefore, "LineUnitAfter": local_LineUnitAfter, "ReadingOrder": local_ReadingOrder, "SpaceBeforeAuto": local_SpaceBeforeAuto, "SpaceAfterAuto": local_SpaceAfterAuto, }
	if (propertyName == "Borders"):
		retVal = this_Range.Borders
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Enable = retVal.Enable
		except:
			local_Enable = None
		try:
			local_DistanceFromTop = retVal.DistanceFromTop
		except:
			local_DistanceFromTop = None
		try:
			local_Shadow = retVal.Shadow
		except:
			local_Shadow = None
		try:
			local_InsideLineStyle = retVal.InsideLineStyle
		except:
			local_InsideLineStyle = None
		try:
			local_OutsideLineStyle = retVal.OutsideLineStyle
		except:
			local_OutsideLineStyle = None
		try:
			local_InsideLineWidth = retVal.InsideLineWidth
		except:
			local_InsideLineWidth = None
		try:
			local_OutsideLineWidth = retVal.OutsideLineWidth
		except:
			local_OutsideLineWidth = None
		try:
			local_InsideColorIndex = retVal.InsideColorIndex
		except:
			local_InsideColorIndex = None
		try:
			local_OutsideColorIndex = retVal.OutsideColorIndex
		except:
			local_OutsideColorIndex = None
		try:
			local_DistanceFromLeft = retVal.DistanceFromLeft
		except:
			local_DistanceFromLeft = None
		try:
			local_DistanceFromBottom = retVal.DistanceFromBottom
		except:
			local_DistanceFromBottom = None
		try:
			local_DistanceFromRight = retVal.DistanceFromRight
		except:
			local_DistanceFromRight = None
		try:
			local_AlwaysInFront = retVal.AlwaysInFront
		except:
			local_AlwaysInFront = None
		try:
			local_SurroundHeader = retVal.SurroundHeader
		except:
			local_SurroundHeader = None
		try:
			local_SurroundFooter = retVal.SurroundFooter
		except:
			local_SurroundFooter = None
		try:
			local_JoinBorders = retVal.JoinBorders
		except:
			local_JoinBorders = None
		try:
			local_HasHorizontal = retVal.HasHorizontal
		except:
			local_HasHorizontal = None
		try:
			local_HasVertical = retVal.HasVertical
		except:
			local_HasVertical = None
		try:
			local_DistanceFrom = retVal.DistanceFrom
		except:
			local_DistanceFrom = None
		try:
			local_EnableFirstPageInSection = retVal.EnableFirstPageInSection
		except:
			local_EnableFirstPageInSection = None
		try:
			local_EnableOtherPagesInSection = retVal.EnableOtherPagesInSection
		except:
			local_EnableOtherPagesInSection = None
		try:
			local_InsideColor = retVal.InsideColor
		except:
			local_InsideColor = None
		try:
			local_OutsideColor = retVal.OutsideColor
		except:
			local_OutsideColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Borders", "Count": local_Count, "Enable": local_Enable, "DistanceFromTop": local_DistanceFromTop, "Shadow": local_Shadow, "InsideLineStyle": local_InsideLineStyle, "OutsideLineStyle": local_OutsideLineStyle, "InsideLineWidth": local_InsideLineWidth, "OutsideLineWidth": local_OutsideLineWidth, "InsideColorIndex": local_InsideColorIndex, "OutsideColorIndex": local_OutsideColorIndex, "DistanceFromLeft": local_DistanceFromLeft, "DistanceFromBottom": local_DistanceFromBottom, "DistanceFromRight": local_DistanceFromRight, "AlwaysInFront": local_AlwaysInFront, "SurroundHeader": local_SurroundHeader, "SurroundFooter": local_SurroundFooter, "JoinBorders": local_JoinBorders, "HasHorizontal": local_HasHorizontal, "HasVertical": local_HasVertical, "DistanceFrom": local_DistanceFrom, "EnableFirstPageInSection": local_EnableFirstPageInSection, "EnableOtherPagesInSection": local_EnableOtherPagesInSection, "InsideColor": local_InsideColor, "OutsideColor": local_OutsideColor, }
	if (propertyName == "Shading"):
		retVal = this_Range.Shading
		try:
			local_ForegroundPatternColorIndex = retVal.ForegroundPatternColorIndex
		except:
			local_ForegroundPatternColorIndex = None
		try:
			local_BackgroundPatternColorIndex = retVal.BackgroundPatternColorIndex
		except:
			local_BackgroundPatternColorIndex = None
		try:
			local_Texture = retVal.Texture
		except:
			local_Texture = None
		try:
			local_ForegroundPatternColor = retVal.ForegroundPatternColor
		except:
			local_ForegroundPatternColor = None
		try:
			local_BackgroundPatternColor = retVal.BackgroundPatternColor
		except:
			local_BackgroundPatternColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Shading", "ForegroundPatternColorIndex": local_ForegroundPatternColorIndex, "BackgroundPatternColorIndex": local_BackgroundPatternColorIndex, "Texture": local_Texture, "ForegroundPatternColor": local_ForegroundPatternColor, "BackgroundPatternColor": local_BackgroundPatternColor, }
	if (propertyName == "TextRetrievalMode"):
		retVal = this_Range.TextRetrievalMode
		try:
			local_ViewType = retVal.ViewType
		except:
			local_ViewType = None
		try:
			local_IncludeHiddenText = retVal.IncludeHiddenText
		except:
			local_IncludeHiddenText = None
		try:
			local_IncludeFieldCodes = retVal.IncludeFieldCodes
		except:
			local_IncludeFieldCodes = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "TextRetrievalMode", "ViewType": local_ViewType, "IncludeHiddenText": local_IncludeHiddenText, "IncludeFieldCodes": local_IncludeFieldCodes, }
	if (propertyName == "Fields"):
		retVal = this_Range.Fields
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Locked = retVal.Locked
		except:
			local_Locked = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Fields", "Count": local_Count, "Locked": local_Locked, }
	if (propertyName == "FormFields"):
		retVal = this_Range.FormFields
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Shaded = retVal.Shaded
		except:
			local_Shaded = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "FormFields", "Count": local_Count, "Shaded": local_Shaded, }
	if (propertyName == "Frames"):
		retVal = this_Range.Frames
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Frames", "Count": local_Count, }
	if (propertyName == "ParagraphFormat"):
		retVal = this_Range.ParagraphFormat
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ParagraphFormat"}
	if (propertyName == "ListFormat"):
		retVal = this_Range.ListFormat
		try:
			local_ListLevelNumber = retVal.ListLevelNumber
		except:
			local_ListLevelNumber = None
		try:
			local_ListValue = retVal.ListValue
		except:
			local_ListValue = None
		try:
			local_SingleList = retVal.SingleList
		except:
			local_SingleList = None
		try:
			local_SingleListTemplate = retVal.SingleListTemplate
		except:
			local_SingleListTemplate = None
		try:
			local_ListType = retVal.ListType
		except:
			local_ListType = None
		try:
			local_ListString = retVal.ListString
		except:
			local_ListString = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ListFormat", "ListLevelNumber": local_ListLevelNumber, "ListValue": local_ListValue, "SingleList": local_SingleList, "SingleListTemplate": local_SingleListTemplate, "ListType": local_ListType, "ListString": local_ListString, }
	if (propertyName == "Bookmarks"):
		retVal = this_Range.Bookmarks
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_DefaultSorting = retVal.DefaultSorting
		except:
			local_DefaultSorting = None
		try:
			local_ShowHidden = retVal.ShowHidden
		except:
			local_ShowHidden = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Bookmarks", "Count": local_Count, "DefaultSorting": local_DefaultSorting, "ShowHidden": local_ShowHidden, }
	if (propertyName == "Bold"):
		retVal = this_Range.Bold
		return retVal
	if (propertyName == "Italic"):
		retVal = this_Range.Italic
		return retVal
	if (propertyName == "Underline"):
		retVal = this_Range.Underline
		return retVal
	if (propertyName == "EmphasisMark"):
		retVal = this_Range.EmphasisMark
		return retVal
	if (propertyName == "DisableCharacterSpaceGrid"):
		retVal = this_Range.DisableCharacterSpaceGrid
		return retVal
	if (propertyName == "Revisions"):
		retVal = this_Range.Revisions
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Revisions", "Count": local_Count, }
	if (propertyName == "Style"):
		retVal = this_Range.Style
		return retVal
	if (propertyName == "StoryLength"):
		retVal = this_Range.StoryLength
		return retVal
	if (propertyName == "LanguageID"):
		retVal = this_Range.LanguageID
		return retVal
	if (propertyName == "SynonymInfo"):
		retVal = this_Range.SynonymInfo
		try:
			local_Word = retVal.Word
		except:
			local_Word = None
		try:
			local_Found = retVal.Found
		except:
			local_Found = None
		try:
			local_MeaningCount = retVal.MeaningCount
		except:
			local_MeaningCount = None
		try:
			local_MeaningList = retVal.MeaningList
		except:
			local_MeaningList = None
		try:
			local_PartOfSpeechList = retVal.PartOfSpeechList
		except:
			local_PartOfSpeechList = None
		try:
			local_SynonymList = retVal.SynonymList
		except:
			local_SynonymList = None
		try:
			local_AntonymList = retVal.AntonymList
		except:
			local_AntonymList = None
		try:
			local_RelatedExpressionList = retVal.RelatedExpressionList
		except:
			local_RelatedExpressionList = None
		try:
			local_RelatedWordList = retVal.RelatedWordList
		except:
			local_RelatedWordList = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "SynonymInfo", "Word": local_Word, "Found": local_Found, "MeaningCount": local_MeaningCount, "MeaningList": local_MeaningList, "PartOfSpeechList": local_PartOfSpeechList, "SynonymList": local_SynonymList, "AntonymList": local_AntonymList, "RelatedExpressionList": local_RelatedExpressionList, "RelatedWordList": local_RelatedWordList, }
	if (propertyName == "Hyperlinks"):
		retVal = this_Range.Hyperlinks
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Hyperlinks", "Count": local_Count, }
	if (propertyName == "ListParagraphs"):
		retVal = this_Range.ListParagraphs
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ListParagraphs", "Count": local_Count, }
	if (propertyName == "Subdocuments"):
		retVal = this_Range.Subdocuments
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Expanded = retVal.Expanded
		except:
			local_Expanded = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Subdocuments", "Count": local_Count, "Expanded": local_Expanded, }
	if (propertyName == "GrammarChecked"):
		retVal = this_Range.GrammarChecked
		return retVal
	if (propertyName == "SpellingChecked"):
		retVal = this_Range.SpellingChecked
		return retVal
	if (propertyName == "HighlightColorIndex"):
		retVal = this_Range.HighlightColorIndex
		return retVal
	if (propertyName == "Columns"):
		retVal = this_Range.Columns
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Columns", "Count": local_Count, "Width": local_Width, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "Rows"):
		retVal = this_Range.Rows
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_AllowBreakAcrossPages = retVal.AllowBreakAcrossPages
		except:
			local_AllowBreakAcrossPages = None
		try:
			local_Alignment = retVal.Alignment
		except:
			local_Alignment = None
		try:
			local_HeadingFormat = retVal.HeadingFormat
		except:
			local_HeadingFormat = None
		try:
			local_SpaceBetweenColumns = retVal.SpaceBetweenColumns
		except:
			local_SpaceBetweenColumns = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_LeftIndent = retVal.LeftIndent
		except:
			local_LeftIndent = None
		try:
			local_WrapAroundText = retVal.WrapAroundText
		except:
			local_WrapAroundText = None
		try:
			local_DistanceTop = retVal.DistanceTop
		except:
			local_DistanceTop = None
		try:
			local_DistanceBottom = retVal.DistanceBottom
		except:
			local_DistanceBottom = None
		try:
			local_DistanceLeft = retVal.DistanceLeft
		except:
			local_DistanceLeft = None
		try:
			local_DistanceRight = retVal.DistanceRight
		except:
			local_DistanceRight = None
		try:
			local_HorizontalPosition = retVal.HorizontalPosition
		except:
			local_HorizontalPosition = None
		try:
			local_VerticalPosition = retVal.VerticalPosition
		except:
			local_VerticalPosition = None
		try:
			local_RelativeHorizontalPosition = retVal.RelativeHorizontalPosition
		except:
			local_RelativeHorizontalPosition = None
		try:
			local_RelativeVerticalPosition = retVal.RelativeVerticalPosition
		except:
			local_RelativeVerticalPosition = None
		try:
			local_AllowOverlap = retVal.AllowOverlap
		except:
			local_AllowOverlap = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_TableDirection = retVal.TableDirection
		except:
			local_TableDirection = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Rows", "Count": local_Count, "AllowBreakAcrossPages": local_AllowBreakAcrossPages, "Alignment": local_Alignment, "HeadingFormat": local_HeadingFormat, "SpaceBetweenColumns": local_SpaceBetweenColumns, "Height": local_Height, "HeightRule": local_HeightRule, "LeftIndent": local_LeftIndent, "WrapAroundText": local_WrapAroundText, "DistanceTop": local_DistanceTop, "DistanceBottom": local_DistanceBottom, "DistanceLeft": local_DistanceLeft, "DistanceRight": local_DistanceRight, "HorizontalPosition": local_HorizontalPosition, "VerticalPosition": local_VerticalPosition, "RelativeHorizontalPosition": local_RelativeHorizontalPosition, "RelativeVerticalPosition": local_RelativeVerticalPosition, "AllowOverlap": local_AllowOverlap, "NestingLevel": local_NestingLevel, "TableDirection": local_TableDirection, }
	if (propertyName == "CanEdit"):
		retVal = this_Range.CanEdit
		return retVal
	if (propertyName == "CanPaste"):
		retVal = this_Range.CanPaste
		return retVal
	if (propertyName == "IsEndOfRowMark"):
		retVal = this_Range.IsEndOfRowMark
		return retVal
	if (propertyName == "BookmarkID"):
		retVal = this_Range.BookmarkID
		return retVal
	if (propertyName == "PreviousBookmarkID"):
		retVal = this_Range.PreviousBookmarkID
		return retVal
	if (propertyName == "Find"):
		retVal = this_Range.Find
		try:
			local_Forward = retVal.Forward
		except:
			local_Forward = None
		try:
			local_Found = retVal.Found
		except:
			local_Found = None
		try:
			local_MatchAllWordForms = retVal.MatchAllWordForms
		except:
			local_MatchAllWordForms = None
		try:
			local_MatchCase = retVal.MatchCase
		except:
			local_MatchCase = None
		try:
			local_MatchWildcards = retVal.MatchWildcards
		except:
			local_MatchWildcards = None
		try:
			local_MatchSoundsLike = retVal.MatchSoundsLike
		except:
			local_MatchSoundsLike = None
		try:
			local_MatchWholeWord = retVal.MatchWholeWord
		except:
			local_MatchWholeWord = None
		try:
			local_MatchFuzzy = retVal.MatchFuzzy
		except:
			local_MatchFuzzy = None
		try:
			local_MatchByte = retVal.MatchByte
		except:
			local_MatchByte = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_LanguageID = retVal.LanguageID
		except:
			local_LanguageID = None
		try:
			local_Highlight = retVal.Highlight
		except:
			local_Highlight = None
		try:
			local_Wrap = retVal.Wrap
		except:
			local_Wrap = None
		try:
			local_Format = retVal.Format
		except:
			local_Format = None
		try:
			local_LanguageIDFarEast = retVal.LanguageIDFarEast
		except:
			local_LanguageIDFarEast = None
		try:
			local_LanguageIDOther = retVal.LanguageIDOther
		except:
			local_LanguageIDOther = None
		try:
			local_CorrectHangulEndings = retVal.CorrectHangulEndings
		except:
			local_CorrectHangulEndings = None
		try:
			local_NoProofing = retVal.NoProofing
		except:
			local_NoProofing = None
		try:
			local_MatchKashida = retVal.MatchKashida
		except:
			local_MatchKashida = None
		try:
			local_MatchDiacritics = retVal.MatchDiacritics
		except:
			local_MatchDiacritics = None
		try:
			local_MatchAlefHamza = retVal.MatchAlefHamza
		except:
			local_MatchAlefHamza = None
		try:
			local_MatchControl = retVal.MatchControl
		except:
			local_MatchControl = None
		try:
			local_MatchPhrase = retVal.MatchPhrase
		except:
			local_MatchPhrase = None
		try:
			local_MatchPrefix = retVal.MatchPrefix
		except:
			local_MatchPrefix = None
		try:
			local_MatchSuffix = retVal.MatchSuffix
		except:
			local_MatchSuffix = None
		try:
			local_IgnoreSpace = retVal.IgnoreSpace
		except:
			local_IgnoreSpace = None
		try:
			local_IgnorePunct = retVal.IgnorePunct
		except:
			local_IgnorePunct = None
		try:
			local_HanjaPhoneticHangul = retVal.HanjaPhoneticHangul
		except:
			local_HanjaPhoneticHangul = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Find", "Forward": local_Forward, "Found": local_Found, "MatchAllWordForms": local_MatchAllWordForms, "MatchCase": local_MatchCase, "MatchWildcards": local_MatchWildcards, "MatchSoundsLike": local_MatchSoundsLike, "MatchWholeWord": local_MatchWholeWord, "MatchFuzzy": local_MatchFuzzy, "MatchByte": local_MatchByte, "Style": local_Style, "Text": local_Text, "LanguageID": local_LanguageID, "Highlight": local_Highlight, "Wrap": local_Wrap, "Format": local_Format, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "CorrectHangulEndings": local_CorrectHangulEndings, "NoProofing": local_NoProofing, "MatchKashida": local_MatchKashida, "MatchDiacritics": local_MatchDiacritics, "MatchAlefHamza": local_MatchAlefHamza, "MatchControl": local_MatchControl, "MatchPhrase": local_MatchPhrase, "MatchPrefix": local_MatchPrefix, "MatchSuffix": local_MatchSuffix, "IgnoreSpace": local_IgnoreSpace, "IgnorePunct": local_IgnorePunct, "HanjaPhoneticHangul": local_HanjaPhoneticHangul, }
	if (propertyName == "PageSetup"):
		retVal = this_Range.PageSetup
		try:
			local_TopMargin = retVal.TopMargin
		except:
			local_TopMargin = None
		try:
			local_BottomMargin = retVal.BottomMargin
		except:
			local_BottomMargin = None
		try:
			local_LeftMargin = retVal.LeftMargin
		except:
			local_LeftMargin = None
		try:
			local_RightMargin = retVal.RightMargin
		except:
			local_RightMargin = None
		try:
			local_Gutter = retVal.Gutter
		except:
			local_Gutter = None
		try:
			local_PageWidth = retVal.PageWidth
		except:
			local_PageWidth = None
		try:
			local_PageHeight = retVal.PageHeight
		except:
			local_PageHeight = None
		try:
			local_Orientation = retVal.Orientation
		except:
			local_Orientation = None
		try:
			local_FirstPageTray = retVal.FirstPageTray
		except:
			local_FirstPageTray = None
		try:
			local_OtherPagesTray = retVal.OtherPagesTray
		except:
			local_OtherPagesTray = None
		try:
			local_VerticalAlignment = retVal.VerticalAlignment
		except:
			local_VerticalAlignment = None
		try:
			local_MirrorMargins = retVal.MirrorMargins
		except:
			local_MirrorMargins = None
		try:
			local_HeaderDistance = retVal.HeaderDistance
		except:
			local_HeaderDistance = None
		try:
			local_FooterDistance = retVal.FooterDistance
		except:
			local_FooterDistance = None
		try:
			local_SectionStart = retVal.SectionStart
		except:
			local_SectionStart = None
		try:
			local_OddAndEvenPagesHeaderFooter = retVal.OddAndEvenPagesHeaderFooter
		except:
			local_OddAndEvenPagesHeaderFooter = None
		try:
			local_DifferentFirstPageHeaderFooter = retVal.DifferentFirstPageHeaderFooter
		except:
			local_DifferentFirstPageHeaderFooter = None
		try:
			local_SuppressEndnotes = retVal.SuppressEndnotes
		except:
			local_SuppressEndnotes = None
		try:
			local_PaperSize = retVal.PaperSize
		except:
			local_PaperSize = None
		try:
			local_TwoPagesOnOne = retVal.TwoPagesOnOne
		except:
			local_TwoPagesOnOne = None
		try:
			local_GutterOnTop = retVal.GutterOnTop
		except:
			local_GutterOnTop = None
		try:
			local_CharsLine = retVal.CharsLine
		except:
			local_CharsLine = None
		try:
			local_LinesPage = retVal.LinesPage
		except:
			local_LinesPage = None
		try:
			local_ShowGrid = retVal.ShowGrid
		except:
			local_ShowGrid = None
		try:
			local_GutterStyle = retVal.GutterStyle
		except:
			local_GutterStyle = None
		try:
			local_SectionDirection = retVal.SectionDirection
		except:
			local_SectionDirection = None
		try:
			local_LayoutMode = retVal.LayoutMode
		except:
			local_LayoutMode = None
		try:
			local_GutterPos = retVal.GutterPos
		except:
			local_GutterPos = None
		try:
			local_BookFoldPrinting = retVal.BookFoldPrinting
		except:
			local_BookFoldPrinting = None
		try:
			local_BookFoldRevPrinting = retVal.BookFoldRevPrinting
		except:
			local_BookFoldRevPrinting = None
		try:
			local_BookFoldPrintingSheets = retVal.BookFoldPrintingSheets
		except:
			local_BookFoldPrintingSheets = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "PageSetup", "TopMargin": local_TopMargin, "BottomMargin": local_BottomMargin, "LeftMargin": local_LeftMargin, "RightMargin": local_RightMargin, "Gutter": local_Gutter, "PageWidth": local_PageWidth, "PageHeight": local_PageHeight, "Orientation": local_Orientation, "FirstPageTray": local_FirstPageTray, "OtherPagesTray": local_OtherPagesTray, "VerticalAlignment": local_VerticalAlignment, "MirrorMargins": local_MirrorMargins, "HeaderDistance": local_HeaderDistance, "FooterDistance": local_FooterDistance, "SectionStart": local_SectionStart, "OddAndEvenPagesHeaderFooter": local_OddAndEvenPagesHeaderFooter, "DifferentFirstPageHeaderFooter": local_DifferentFirstPageHeaderFooter, "SuppressEndnotes": local_SuppressEndnotes, "PaperSize": local_PaperSize, "TwoPagesOnOne": local_TwoPagesOnOne, "GutterOnTop": local_GutterOnTop, "CharsLine": local_CharsLine, "LinesPage": local_LinesPage, "ShowGrid": local_ShowGrid, "GutterStyle": local_GutterStyle, "SectionDirection": local_SectionDirection, "LayoutMode": local_LayoutMode, "GutterPos": local_GutterPos, "BookFoldPrinting": local_BookFoldPrinting, "BookFoldRevPrinting": local_BookFoldRevPrinting, "BookFoldPrintingSheets": local_BookFoldPrintingSheets, }
	if (propertyName == "ShapeRange"):
		retVal = this_Range.ShapeRange
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_AutoShapeType = retVal.AutoShapeType
		except:
			local_AutoShapeType = None
		try:
			local_ConnectionSiteCount = retVal.ConnectionSiteCount
		except:
			local_ConnectionSiteCount = None
		try:
			local_Connector = retVal.Connector
		except:
			local_Connector = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HorizontalFlip = retVal.HorizontalFlip
		except:
			local_HorizontalFlip = None
		try:
			local_Left = retVal.Left
		except:
			local_Left = None
		try:
			local_LockAspectRatio = retVal.LockAspectRatio
		except:
			local_LockAspectRatio = None
		try:
			local_Name = retVal.Name
		except:
			local_Name = None
		try:
			local_Rotation = retVal.Rotation
		except:
			local_Rotation = None
		try:
			local_Top = retVal.Top
		except:
			local_Top = None
		try:
			local_Type = retVal.Type
		except:
			local_Type = None
		try:
			local_VerticalFlip = retVal.VerticalFlip
		except:
			local_VerticalFlip = None
		try:
			local_Vertices = retVal.Vertices
		except:
			local_Vertices = None
		try:
			local_Visible = retVal.Visible
		except:
			local_Visible = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_ZOrderPosition = retVal.ZOrderPosition
		except:
			local_ZOrderPosition = None
		try:
			local_RelativeHorizontalPosition = retVal.RelativeHorizontalPosition
		except:
			local_RelativeHorizontalPosition = None
		try:
			local_RelativeVerticalPosition = retVal.RelativeVerticalPosition
		except:
			local_RelativeVerticalPosition = None
		try:
			local_LockAnchor = retVal.LockAnchor
		except:
			local_LockAnchor = None
		try:
			local_AlternativeText = retVal.AlternativeText
		except:
			local_AlternativeText = None
		try:
			local_HasDiagram = retVal.HasDiagram
		except:
			local_HasDiagram = None
		try:
			local_HasDiagramNode = retVal.HasDiagramNode
		except:
			local_HasDiagramNode = None
		try:
			local_Child = retVal.Child
		except:
			local_Child = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_LayoutInCell = retVal.LayoutInCell
		except:
			local_LayoutInCell = None
		try:
			local_LeftRelative = retVal.LeftRelative
		except:
			local_LeftRelative = None
		try:
			local_TopRelative = retVal.TopRelative
		except:
			local_TopRelative = None
		try:
			local_WidthRelative = retVal.WidthRelative
		except:
			local_WidthRelative = None
		try:
			local_HeightRelative = retVal.HeightRelative
		except:
			local_HeightRelative = None
		try:
			local_RelativeHorizontalSize = retVal.RelativeHorizontalSize
		except:
			local_RelativeHorizontalSize = None
		try:
			local_RelativeVerticalSize = retVal.RelativeVerticalSize
		except:
			local_RelativeVerticalSize = None
		try:
			local_ShapeStyle = retVal.ShapeStyle
		except:
			local_ShapeStyle = None
		try:
			local_BackgroundStyle = retVal.BackgroundStyle
		except:
			local_BackgroundStyle = None
		try:
			local_Title = retVal.Title
		except:
			local_Title = None
		try:
			local_GraphicStyle = retVal.GraphicStyle
		except:
			local_GraphicStyle = None
		try:
			local_Decorative = retVal.Decorative
		except:
			local_Decorative = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ShapeRange", "Count": local_Count, "AutoShapeType": local_AutoShapeType, "ConnectionSiteCount": local_ConnectionSiteCount, "Connector": local_Connector, "Height": local_Height, "HorizontalFlip": local_HorizontalFlip, "Left": local_Left, "LockAspectRatio": local_LockAspectRatio, "Name": local_Name, "Rotation": local_Rotation, "Top": local_Top, "Type": local_Type, "VerticalFlip": local_VerticalFlip, "Vertices": local_Vertices, "Visible": local_Visible, "Width": local_Width, "ZOrderPosition": local_ZOrderPosition, "RelativeHorizontalPosition": local_RelativeHorizontalPosition, "RelativeVerticalPosition": local_RelativeVerticalPosition, "LockAnchor": local_LockAnchor, "AlternativeText": local_AlternativeText, "HasDiagram": local_HasDiagram, "HasDiagramNode": local_HasDiagramNode, "Child": local_Child, "ID": local_ID, "LayoutInCell": local_LayoutInCell, "LeftRelative": local_LeftRelative, "TopRelative": local_TopRelative, "WidthRelative": local_WidthRelative, "HeightRelative": local_HeightRelative, "RelativeHorizontalSize": local_RelativeHorizontalSize, "RelativeVerticalSize": local_RelativeVerticalSize, "ShapeStyle": local_ShapeStyle, "BackgroundStyle": local_BackgroundStyle, "Title": local_Title, "GraphicStyle": local_GraphicStyle, "Decorative": local_Decorative, }
	if (propertyName == "Case"):
		retVal = this_Range.Case
		return retVal
	if (propertyName == "ReadabilityStatistics"):
		retVal = this_Range.ReadabilityStatistics
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ReadabilityStatistics", "Count": local_Count, }
	if (propertyName == "GrammaticalErrors"):
		retVal = this_Range.GrammaticalErrors
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Type = retVal.Type
		except:
			local_Type = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ProofreadingErrors", "Count": local_Count, "Type": local_Type, }
	if (propertyName == "SpellingErrors"):
		retVal = this_Range.SpellingErrors
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Type = retVal.Type
		except:
			local_Type = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ProofreadingErrors", "Count": local_Count, "Type": local_Type, }
	if (propertyName == "Orientation"):
		retVal = this_Range.Orientation
		return retVal
	if (propertyName == "InlineShapes"):
		retVal = this_Range.InlineShapes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "InlineShapes", "Count": local_Count, }
	if (propertyName == "NextStoryRange"):
		retVal = this_Range.NextStoryRange
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_Start = retVal.Start
		except:
			local_Start = None
		try:
			local_End = retVal.End
		except:
			local_End = None
		try:
			local_StoryType = retVal.StoryType
		except:
			local_StoryType = None
		try:
			local_Bold = retVal.Bold
		except:
			local_Bold = None
		try:
			local_Italic = retVal.Italic
		except:
			local_Italic = None
		try:
			local_Underline = retVal.Underline
		except:
			local_Underline = None
		try:
			local_EmphasisMark = retVal.EmphasisMark
		except:
			local_EmphasisMark = None
		try:
			local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
		except:
			local_DisableCharacterSpaceGrid = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_StoryLength = retVal.StoryLength
		except:
			local_StoryLength = None
		try:
			local_LanguageID = retVal.LanguageID
		except:
			local_LanguageID = None
		try:
			local_GrammarChecked = retVal.GrammarChecked
		except:
			local_GrammarChecked = None
		try:
			local_SpellingChecked = retVal.SpellingChecked
		except:
			local_SpellingChecked = None
		try:
			local_HighlightColorIndex = retVal.HighlightColorIndex
		except:
			local_HighlightColorIndex = None
		try:
			local_CanEdit = retVal.CanEdit
		except:
			local_CanEdit = None
		try:
			local_CanPaste = retVal.CanPaste
		except:
			local_CanPaste = None
		try:
			local_IsEndOfRowMark = retVal.IsEndOfRowMark
		except:
			local_IsEndOfRowMark = None
		try:
			local_BookmarkID = retVal.BookmarkID
		except:
			local_BookmarkID = None
		try:
			local_PreviousBookmarkID = retVal.PreviousBookmarkID
		except:
			local_PreviousBookmarkID = None
		try:
			local_Case = retVal.Case
		except:
			local_Case = None
		try:
			local_Information = retVal.Information
		except:
			local_Information = None
		try:
			local_Orientation = retVal.Orientation
		except:
			local_Orientation = None
		try:
			local_LanguageIDFarEast = retVal.LanguageIDFarEast
		except:
			local_LanguageIDFarEast = None
		try:
			local_LanguageIDOther = retVal.LanguageIDOther
		except:
			local_LanguageIDOther = None
		try:
			local_LanguageDetected = retVal.LanguageDetected
		except:
			local_LanguageDetected = None
		try:
			local_FitTextWidth = retVal.FitTextWidth
		except:
			local_FitTextWidth = None
		try:
			local_HorizontalInVertical = retVal.HorizontalInVertical
		except:
			local_HorizontalInVertical = None
		try:
			local_TwoLinesInOne = retVal.TwoLinesInOne
		except:
			local_TwoLinesInOne = None
		try:
			local_CombineCharacters = retVal.CombineCharacters
		except:
			local_CombineCharacters = None
		try:
			local_NoProofing = retVal.NoProofing
		except:
			local_NoProofing = None
		try:
			local_CharacterWidth = retVal.CharacterWidth
		except:
			local_CharacterWidth = None
		try:
			local_Kana = retVal.Kana
		except:
			local_Kana = None
		try:
			local_BoldBi = retVal.BoldBi
		except:
			local_BoldBi = None
		try:
			local_ItalicBi = retVal.ItalicBi
		except:
			local_ItalicBi = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_ShowAll = retVal.ShowAll
		except:
			local_ShowAll = None
		try:
			local_CharacterStyle = retVal.CharacterStyle
		except:
			local_CharacterStyle = None
		try:
			local_ParagraphStyle = retVal.ParagraphStyle
		except:
			local_ParagraphStyle = None
		try:
			local_ListStyle = retVal.ListStyle
		except:
			local_ListStyle = None
		try:
			local_TableStyle = retVal.TableStyle
		except:
			local_TableStyle = None
		try:
			local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
		except:
			local_TextVisibleOnScreen = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }
	if (propertyName == "LanguageIDFarEast"):
		retVal = this_Range.LanguageIDFarEast
		return retVal
	if (propertyName == "LanguageIDOther"):
		retVal = this_Range.LanguageIDOther
		return retVal
	if (propertyName == "LanguageDetected"):
		retVal = this_Range.LanguageDetected
		return retVal
	if (propertyName == "FitTextWidth"):
		retVal = this_Range.FitTextWidth
		return retVal
	if (propertyName == "HorizontalInVertical"):
		retVal = this_Range.HorizontalInVertical
		return retVal
	if (propertyName == "TwoLinesInOne"):
		retVal = this_Range.TwoLinesInOne
		return retVal
	if (propertyName == "CombineCharacters"):
		retVal = this_Range.CombineCharacters
		return retVal
	if (propertyName == "NoProofing"):
		retVal = this_Range.NoProofing
		return retVal
	if (propertyName == "TopLevelTables"):
		retVal = this_Range.TopLevelTables
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Tables", "Count": local_Count, "NestingLevel": local_NestingLevel, }
	if (propertyName == "Scripts"):
		retVal = this_Range.Scripts
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Scripts"}
	if (propertyName == "CharacterWidth"):
		retVal = this_Range.CharacterWidth
		return retVal
	if (propertyName == "Kana"):
		retVal = this_Range.Kana
		return retVal
	if (propertyName == "BoldBi"):
		retVal = this_Range.BoldBi
		return retVal
	if (propertyName == "ItalicBi"):
		retVal = this_Range.ItalicBi
		return retVal
	if (propertyName == "ID"):
		retVal = this_Range.ID
		return retVal
	if (propertyName == "HTMLDivisions"):
		retVal = this_Range.HTMLDivisions
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "HTMLDivisions", "Count": local_Count, "NestingLevel": local_NestingLevel, }
	if (propertyName == "SmartTags"):
		retVal = this_Range.SmartTags
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "SmartTags", "Count": local_Count, }
	if (propertyName == "ShowAll"):
		retVal = this_Range.ShowAll
		return retVal
	if (propertyName == "Document"):
		retVal = this_Range.Document
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Document"}
	if (propertyName == "FootnoteOptions"):
		retVal = this_Range.FootnoteOptions
		try:
			local_Location = retVal.Location
		except:
			local_Location = None
		try:
			local_NumberStyle = retVal.NumberStyle
		except:
			local_NumberStyle = None
		try:
			local_StartingNumber = retVal.StartingNumber
		except:
			local_StartingNumber = None
		try:
			local_NumberingRule = retVal.NumberingRule
		except:
			local_NumberingRule = None
		try:
			local_LayoutColumns = retVal.LayoutColumns
		except:
			local_LayoutColumns = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "FootnoteOptions", "Location": local_Location, "NumberStyle": local_NumberStyle, "StartingNumber": local_StartingNumber, "NumberingRule": local_NumberingRule, "LayoutColumns": local_LayoutColumns, }
	if (propertyName == "EndnoteOptions"):
		retVal = this_Range.EndnoteOptions
		try:
			local_Location = retVal.Location
		except:
			local_Location = None
		try:
			local_NumberStyle = retVal.NumberStyle
		except:
			local_NumberStyle = None
		try:
			local_StartingNumber = retVal.StartingNumber
		except:
			local_StartingNumber = None
		try:
			local_NumberingRule = retVal.NumberingRule
		except:
			local_NumberingRule = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "EndnoteOptions", "Location": local_Location, "NumberStyle": local_NumberStyle, "StartingNumber": local_StartingNumber, "NumberingRule": local_NumberingRule, }
	if (propertyName == "XMLNodes"):
		retVal = this_Range.XMLNodes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "XMLNodes", "Count": local_Count, }
	if (propertyName == "XMLParentNode"):
		retVal = this_Range.XMLParentNode
		try:
			local_BaseName = retVal.BaseName
		except:
			local_BaseName = None
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_NamespaceURI = retVal.NamespaceURI
		except:
			local_NamespaceURI = None
		try:
			local_NodeType = retVal.NodeType
		except:
			local_NodeType = None
		try:
			local_NodeValue = retVal.NodeValue
		except:
			local_NodeValue = None
		try:
			local_HasChildNodes = retVal.HasChildNodes
		except:
			local_HasChildNodes = None
		try:
			local_Level = retVal.Level
		except:
			local_Level = None
		try:
			local_ValidationStatus = retVal.ValidationStatus
		except:
			local_ValidationStatus = None
		try:
			local_ValidationErrorText = retVal.ValidationErrorText
		except:
			local_ValidationErrorText = None
		try:
			local_PlaceholderText = retVal.PlaceholderText
		except:
			local_PlaceholderText = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "XMLNode", "BaseName": local_BaseName, "Text": local_Text, "NamespaceURI": local_NamespaceURI, "NodeType": local_NodeType, "NodeValue": local_NodeValue, "HasChildNodes": local_HasChildNodes, "Level": local_Level, "ValidationStatus": local_ValidationStatus, "ValidationErrorText": local_ValidationErrorText, "PlaceholderText": local_PlaceholderText, }
	if (propertyName == "Editors"):
		retVal = this_Range.Editors
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Editors", "Count": local_Count, }
	if (propertyName == "EnhMetaFileBits"):
		retVal = this_Range.EnhMetaFileBits
		return retVal
	if (propertyName == "OMaths"):
		retVal = this_Range.OMaths
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "OMaths", "Count": local_Count, }
	if (propertyName == "CharacterStyle"):
		retVal = this_Range.CharacterStyle
		return retVal
	if (propertyName == "ParagraphStyle"):
		retVal = this_Range.ParagraphStyle
		return retVal
	if (propertyName == "ListStyle"):
		retVal = this_Range.ListStyle
		return retVal
	if (propertyName == "TableStyle"):
		retVal = this_Range.TableStyle
		return retVal
	if (propertyName == "ContentControls"):
		retVal = this_Range.ContentControls
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ContentControls", "Count": local_Count, }
	if (propertyName == "WordOpenXML"):
		retVal = this_Range.WordOpenXML
		return retVal
	if (propertyName == "ParentContentControl"):
		retVal = this_Range.ParentContentControl
		try:
			local_LockContentControl = retVal.LockContentControl
		except:
			local_LockContentControl = None
		try:
			local_LockContents = retVal.LockContents
		except:
			local_LockContents = None
		try:
			local_Type = retVal.Type
		except:
			local_Type = None
		try:
			local_Title = retVal.Title
		except:
			local_Title = None
		try:
			local_DateDisplayFormat = retVal.DateDisplayFormat
		except:
			local_DateDisplayFormat = None
		try:
			local_MultiLine = retVal.MultiLine
		except:
			local_MultiLine = None
		try:
			local_Temporary = retVal.Temporary
		except:
			local_Temporary = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_ShowingPlaceholderText = retVal.ShowingPlaceholderText
		except:
			local_ShowingPlaceholderText = None
		try:
			local_DateStorageFormat = retVal.DateStorageFormat
		except:
			local_DateStorageFormat = None
		try:
			local_BuildingBlockType = retVal.BuildingBlockType
		except:
			local_BuildingBlockType = None
		try:
			local_BuildingBlockCategory = retVal.BuildingBlockCategory
		except:
			local_BuildingBlockCategory = None
		try:
			local_DateDisplayLocale = retVal.DateDisplayLocale
		except:
			local_DateDisplayLocale = None
		try:
			local_DefaultTextStyle = retVal.DefaultTextStyle
		except:
			local_DefaultTextStyle = None
		try:
			local_DateCalendarType = retVal.DateCalendarType
		except:
			local_DateCalendarType = None
		try:
			local_Tag = retVal.Tag
		except:
			local_Tag = None
		try:
			local_Checked = retVal.Checked
		except:
			local_Checked = None
		try:
			local_Color = retVal.Color
		except:
			local_Color = None
		try:
			local_Appearance = retVal.Appearance
		except:
			local_Appearance = None
		try:
			local_Level = retVal.Level
		except:
			local_Level = None
		try:
			local_RepeatingSectionItemTitle = retVal.RepeatingSectionItemTitle
		except:
			local_RepeatingSectionItemTitle = None
		try:
			local_AllowInsertDeleteSection = retVal.AllowInsertDeleteSection
		except:
			local_AllowInsertDeleteSection = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ContentControl", "LockContentControl": local_LockContentControl, "LockContents": local_LockContents, "Type": local_Type, "Title": local_Title, "DateDisplayFormat": local_DateDisplayFormat, "MultiLine": local_MultiLine, "Temporary": local_Temporary, "ID": local_ID, "ShowingPlaceholderText": local_ShowingPlaceholderText, "DateStorageFormat": local_DateStorageFormat, "BuildingBlockType": local_BuildingBlockType, "BuildingBlockCategory": local_BuildingBlockCategory, "DateDisplayLocale": local_DateDisplayLocale, "DefaultTextStyle": local_DefaultTextStyle, "DateCalendarType": local_DateCalendarType, "Tag": local_Tag, "Checked": local_Checked, "Color": local_Color, "Appearance": local_Appearance, "Level": local_Level, "RepeatingSectionItemTitle": local_RepeatingSectionItemTitle, "AllowInsertDeleteSection": local_AllowInsertDeleteSection, }
	if (propertyName == "Locks"):
		retVal = this_Range.Locks
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "CoAuthLocks", "Count": local_Count, }
	if (propertyName == "Updates"):
		retVal = this_Range.Updates
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "CoAuthUpdates", "Count": local_Count, }
	if (propertyName == "Conflicts"):
		retVal = this_Range.Conflicts
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Conflicts", "Count": local_Count, }
	if (propertyName == "TextVisibleOnScreen"):
		retVal = this_Range.TextVisibleOnScreen
		return retVal


# Tool: 297
@mcp.tool()
async def word_Range_set_Property(this_Range_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Range
	
	propertyName: Name of the property. Can be one of ...
		Text, FormattedText, Start, End, Font, Borders, TextRetrievalMode, ParagraphFormat, Bold, Italic, Underline, EmphasisMark, DisableCharacterSpaceGrid, Style, LanguageID, GrammarChecked, SpellingChecked, HighlightColorIndex, PageSetup, Case, Orientation, LanguageIDFarEast, LanguageIDOther, LanguageDetected, FitTextWidth, HorizontalInVertical, TwoLinesInOne, CombineCharacters, NoProofing, CharacterWidth, Kana, BoldBi, ItalicBi, ID, ShowAll
	"""
	this_Range = get_object(this_Range_wordObjId)
	
	EnsureWord()
	if (propertyName == "Text"):
		this_Range.Text = propertyValue
	if (propertyName == "FormattedText"):
		this_Range.FormattedText = propertyValue
	if (propertyName == "Start"):
		this_Range.Start = propertyValue
	if (propertyName == "End"):
		this_Range.End = propertyValue
	if (propertyName == "Font"):
		this_Range.Font = propertyValue
	if (propertyName == "Borders"):
		this_Range.Borders = propertyValue
	if (propertyName == "TextRetrievalMode"):
		this_Range.TextRetrievalMode = propertyValue
	if (propertyName == "ParagraphFormat"):
		this_Range.ParagraphFormat = propertyValue
	if (propertyName == "Bold"):
		this_Range.Bold = propertyValue
	if (propertyName == "Italic"):
		this_Range.Italic = propertyValue
	if (propertyName == "Underline"):
		this_Range.Underline = propertyValue
	if (propertyName == "EmphasisMark"):
		this_Range.EmphasisMark = propertyValue
	if (propertyName == "DisableCharacterSpaceGrid"):
		this_Range.DisableCharacterSpaceGrid = propertyValue
	if (propertyName == "Style"):
		this_Range.Style = propertyValue
	if (propertyName == "LanguageID"):
		this_Range.LanguageID = propertyValue
	if (propertyName == "GrammarChecked"):
		this_Range.GrammarChecked = propertyValue
	if (propertyName == "SpellingChecked"):
		this_Range.SpellingChecked = propertyValue
	if (propertyName == "HighlightColorIndex"):
		this_Range.HighlightColorIndex = propertyValue
	if (propertyName == "PageSetup"):
		this_Range.PageSetup = propertyValue
	if (propertyName == "Case"):
		this_Range.Case = propertyValue
	if (propertyName == "Orientation"):
		this_Range.Orientation = propertyValue
	if (propertyName == "LanguageIDFarEast"):
		this_Range.LanguageIDFarEast = propertyValue
	if (propertyName == "LanguageIDOther"):
		this_Range.LanguageIDOther = propertyValue
	if (propertyName == "LanguageDetected"):
		this_Range.LanguageDetected = propertyValue
	if (propertyName == "FitTextWidth"):
		this_Range.FitTextWidth = propertyValue
	if (propertyName == "HorizontalInVertical"):
		this_Range.HorizontalInVertical = propertyValue
	if (propertyName == "TwoLinesInOne"):
		this_Range.TwoLinesInOne = propertyValue
	if (propertyName == "CombineCharacters"):
		this_Range.CombineCharacters = propertyValue
	if (propertyName == "NoProofing"):
		this_Range.NoProofing = propertyValue
	if (propertyName == "CharacterWidth"):
		this_Range.CharacterWidth = propertyValue
	if (propertyName == "Kana"):
		this_Range.Kana = propertyValue
	if (propertyName == "BoldBi"):
		this_Range.BoldBi = propertyValue
	if (propertyName == "ItalicBi"):
		this_Range.ItalicBi = propertyValue
	if (propertyName == "ID"):
		this_Range.ID = propertyValue
	if (propertyName == "ShowAll"):
		this_Range.ShowAll = propertyValue


# Tool: 298
@mcp.tool()
async def word_Paragraph_CloseUp(this_Paragraph_wordObjId: str):
	"""This tool calls the CloseUp methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.CloseUp()


# Tool: 299
@mcp.tool()
async def word_Paragraph_OpenUp(this_Paragraph_wordObjId: str):
	"""This tool calls the OpenUp methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.OpenUp()


# Tool: 300
@mcp.tool()
async def word_Paragraph_OpenOrCloseUp(this_Paragraph_wordObjId: str):
	"""This tool calls the OpenOrCloseUp methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.OpenOrCloseUp()


# Tool: 301
@mcp.tool()
async def word_Paragraph_TabHangingIndent(this_Paragraph_wordObjId: str, Count: int):
	"""This tool calls the TabHangingIndent methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
	
	Parameters:
		Count: the Count as int
	"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.TabHangingIndent(Count)


# Tool: 302
@mcp.tool()
async def word_Paragraph_TabIndent(this_Paragraph_wordObjId: str, Count: int):
	"""This tool calls the TabIndent methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
	
	Parameters:
		Count: the Count as int
	"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.TabIndent(Count)


# Tool: 303
@mcp.tool()
async def word_Paragraph_Reset(this_Paragraph_wordObjId: str):
	"""This tool calls the Reset methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.Reset()


# Tool: 304
@mcp.tool()
async def word_Paragraph_Space1(this_Paragraph_wordObjId: str):
	"""This tool calls the Space1 methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.Space1()


# Tool: 305
@mcp.tool()
async def word_Paragraph_Space15(this_Paragraph_wordObjId: str):
	"""This tool calls the Space15 methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.Space15()


# Tool: 306
@mcp.tool()
async def word_Paragraph_Space2(this_Paragraph_wordObjId: str):
	"""This tool calls the Space2 methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.Space2()


# Tool: 307
@mcp.tool()
async def word_Paragraph_IndentCharWidth(this_Paragraph_wordObjId: str, Count: int):
	"""This tool calls the IndentCharWidth methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
	
	Parameters:
		Count: the Count as int
	"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.IndentCharWidth(Count)


# Tool: 308
@mcp.tool()
async def word_Paragraph_IndentFirstLineCharWidth(this_Paragraph_wordObjId: str, Count: int):
	"""This tool calls the IndentFirstLineCharWidth methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
	
	Parameters:
		Count: the Count as int
	"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.IndentFirstLineCharWidth(Count)


# Tool: 309
@mcp.tool()
async def word_Paragraph_Next(this_Paragraph_wordObjId: str, Count):
	"""This tool calls the Next methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
	
	Parameters:
		Count: the Count as VT_VARIANT
	"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	Count = tryParseString(Count)
	retVal = this_Paragraph.Next(Count)
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_Alignment = retVal.Alignment
	except:
		local_Alignment = None
	try:
		local_KeepTogether = retVal.KeepTogether
	except:
		local_KeepTogether = None
	try:
		local_KeepWithNext = retVal.KeepWithNext
	except:
		local_KeepWithNext = None
	try:
		local_PageBreakBefore = retVal.PageBreakBefore
	except:
		local_PageBreakBefore = None
	try:
		local_NoLineNumber = retVal.NoLineNumber
	except:
		local_NoLineNumber = None
	try:
		local_RightIndent = retVal.RightIndent
	except:
		local_RightIndent = None
	try:
		local_LeftIndent = retVal.LeftIndent
	except:
		local_LeftIndent = None
	try:
		local_FirstLineIndent = retVal.FirstLineIndent
	except:
		local_FirstLineIndent = None
	try:
		local_LineSpacing = retVal.LineSpacing
	except:
		local_LineSpacing = None
	try:
		local_LineSpacingRule = retVal.LineSpacingRule
	except:
		local_LineSpacingRule = None
	try:
		local_SpaceBefore = retVal.SpaceBefore
	except:
		local_SpaceBefore = None
	try:
		local_SpaceAfter = retVal.SpaceAfter
	except:
		local_SpaceAfter = None
	try:
		local_Hyphenation = retVal.Hyphenation
	except:
		local_Hyphenation = None
	try:
		local_WidowControl = retVal.WidowControl
	except:
		local_WidowControl = None
	try:
		local_FarEastLineBreakControl = retVal.FarEastLineBreakControl
	except:
		local_FarEastLineBreakControl = None
	try:
		local_WordWrap = retVal.WordWrap
	except:
		local_WordWrap = None
	try:
		local_HangingPunctuation = retVal.HangingPunctuation
	except:
		local_HangingPunctuation = None
	try:
		local_HalfWidthPunctuationOnTopOfLine = retVal.HalfWidthPunctuationOnTopOfLine
	except:
		local_HalfWidthPunctuationOnTopOfLine = None
	try:
		local_AddSpaceBetweenFarEastAndAlpha = retVal.AddSpaceBetweenFarEastAndAlpha
	except:
		local_AddSpaceBetweenFarEastAndAlpha = None
	try:
		local_AddSpaceBetweenFarEastAndDigit = retVal.AddSpaceBetweenFarEastAndDigit
	except:
		local_AddSpaceBetweenFarEastAndDigit = None
	try:
		local_BaseLineAlignment = retVal.BaseLineAlignment
	except:
		local_BaseLineAlignment = None
	try:
		local_AutoAdjustRightIndent = retVal.AutoAdjustRightIndent
	except:
		local_AutoAdjustRightIndent = None
	try:
		local_DisableLineHeightGrid = retVal.DisableLineHeightGrid
	except:
		local_DisableLineHeightGrid = None
	try:
		local_OutlineLevel = retVal.OutlineLevel
	except:
		local_OutlineLevel = None
	try:
		local_CharacterUnitRightIndent = retVal.CharacterUnitRightIndent
	except:
		local_CharacterUnitRightIndent = None
	try:
		local_CharacterUnitLeftIndent = retVal.CharacterUnitLeftIndent
	except:
		local_CharacterUnitLeftIndent = None
	try:
		local_CharacterUnitFirstLineIndent = retVal.CharacterUnitFirstLineIndent
	except:
		local_CharacterUnitFirstLineIndent = None
	try:
		local_LineUnitBefore = retVal.LineUnitBefore
	except:
		local_LineUnitBefore = None
	try:
		local_LineUnitAfter = retVal.LineUnitAfter
	except:
		local_LineUnitAfter = None
	try:
		local_ReadingOrder = retVal.ReadingOrder
	except:
		local_ReadingOrder = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_SpaceBeforeAuto = retVal.SpaceBeforeAuto
	except:
		local_SpaceBeforeAuto = None
	try:
		local_SpaceAfterAuto = retVal.SpaceAfterAuto
	except:
		local_SpaceAfterAuto = None
	try:
		local_IsStyleSeparator = retVal.IsStyleSeparator
	except:
		local_IsStyleSeparator = None
	try:
		local_MirrorIndents = retVal.MirrorIndents
	except:
		local_MirrorIndents = None
	try:
		local_TextboxTightWrap = retVal.TextboxTightWrap
	except:
		local_TextboxTightWrap = None
	try:
		local_ListNumberOriginal = retVal.ListNumberOriginal
	except:
		local_ListNumberOriginal = None
	try:
		local_ParaID = retVal.ParaID
	except:
		local_ParaID = None
	try:
		local_TextID = retVal.TextID
	except:
		local_TextID = None
	try:
		local_CollapsedState = retVal.CollapsedState
	except:
		local_CollapsedState = None
	try:
		local_CollapseHeadingByDefault = retVal.CollapseHeadingByDefault
	except:
		local_CollapseHeadingByDefault = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Paragraph", "Style": local_Style, "Alignment": local_Alignment, "KeepTogether": local_KeepTogether, "KeepWithNext": local_KeepWithNext, "PageBreakBefore": local_PageBreakBefore, "NoLineNumber": local_NoLineNumber, "RightIndent": local_RightIndent, "LeftIndent": local_LeftIndent, "FirstLineIndent": local_FirstLineIndent, "LineSpacing": local_LineSpacing, "LineSpacingRule": local_LineSpacingRule, "SpaceBefore": local_SpaceBefore, "SpaceAfter": local_SpaceAfter, "Hyphenation": local_Hyphenation, "WidowControl": local_WidowControl, "FarEastLineBreakControl": local_FarEastLineBreakControl, "WordWrap": local_WordWrap, "HangingPunctuation": local_HangingPunctuation, "HalfWidthPunctuationOnTopOfLine": local_HalfWidthPunctuationOnTopOfLine, "AddSpaceBetweenFarEastAndAlpha": local_AddSpaceBetweenFarEastAndAlpha, "AddSpaceBetweenFarEastAndDigit": local_AddSpaceBetweenFarEastAndDigit, "BaseLineAlignment": local_BaseLineAlignment, "AutoAdjustRightIndent": local_AutoAdjustRightIndent, "DisableLineHeightGrid": local_DisableLineHeightGrid, "OutlineLevel": local_OutlineLevel, "CharacterUnitRightIndent": local_CharacterUnitRightIndent, "CharacterUnitLeftIndent": local_CharacterUnitLeftIndent, "CharacterUnitFirstLineIndent": local_CharacterUnitFirstLineIndent, "LineUnitBefore": local_LineUnitBefore, "LineUnitAfter": local_LineUnitAfter, "ReadingOrder": local_ReadingOrder, "ID": local_ID, "SpaceBeforeAuto": local_SpaceBeforeAuto, "SpaceAfterAuto": local_SpaceAfterAuto, "IsStyleSeparator": local_IsStyleSeparator, "MirrorIndents": local_MirrorIndents, "TextboxTightWrap": local_TextboxTightWrap, "ListNumberOriginal": local_ListNumberOriginal, "ParaID": local_ParaID, "TextID": local_TextID, "CollapsedState": local_CollapsedState, "CollapseHeadingByDefault": local_CollapseHeadingByDefault, }


# Tool: 310
@mcp.tool()
async def word_Paragraph_Previous(this_Paragraph_wordObjId: str, Count):
	"""This tool calls the Previous methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
	
	Parameters:
		Count: the Count as VT_VARIANT
	"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	Count = tryParseString(Count)
	retVal = this_Paragraph.Previous(Count)
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_Alignment = retVal.Alignment
	except:
		local_Alignment = None
	try:
		local_KeepTogether = retVal.KeepTogether
	except:
		local_KeepTogether = None
	try:
		local_KeepWithNext = retVal.KeepWithNext
	except:
		local_KeepWithNext = None
	try:
		local_PageBreakBefore = retVal.PageBreakBefore
	except:
		local_PageBreakBefore = None
	try:
		local_NoLineNumber = retVal.NoLineNumber
	except:
		local_NoLineNumber = None
	try:
		local_RightIndent = retVal.RightIndent
	except:
		local_RightIndent = None
	try:
		local_LeftIndent = retVal.LeftIndent
	except:
		local_LeftIndent = None
	try:
		local_FirstLineIndent = retVal.FirstLineIndent
	except:
		local_FirstLineIndent = None
	try:
		local_LineSpacing = retVal.LineSpacing
	except:
		local_LineSpacing = None
	try:
		local_LineSpacingRule = retVal.LineSpacingRule
	except:
		local_LineSpacingRule = None
	try:
		local_SpaceBefore = retVal.SpaceBefore
	except:
		local_SpaceBefore = None
	try:
		local_SpaceAfter = retVal.SpaceAfter
	except:
		local_SpaceAfter = None
	try:
		local_Hyphenation = retVal.Hyphenation
	except:
		local_Hyphenation = None
	try:
		local_WidowControl = retVal.WidowControl
	except:
		local_WidowControl = None
	try:
		local_FarEastLineBreakControl = retVal.FarEastLineBreakControl
	except:
		local_FarEastLineBreakControl = None
	try:
		local_WordWrap = retVal.WordWrap
	except:
		local_WordWrap = None
	try:
		local_HangingPunctuation = retVal.HangingPunctuation
	except:
		local_HangingPunctuation = None
	try:
		local_HalfWidthPunctuationOnTopOfLine = retVal.HalfWidthPunctuationOnTopOfLine
	except:
		local_HalfWidthPunctuationOnTopOfLine = None
	try:
		local_AddSpaceBetweenFarEastAndAlpha = retVal.AddSpaceBetweenFarEastAndAlpha
	except:
		local_AddSpaceBetweenFarEastAndAlpha = None
	try:
		local_AddSpaceBetweenFarEastAndDigit = retVal.AddSpaceBetweenFarEastAndDigit
	except:
		local_AddSpaceBetweenFarEastAndDigit = None
	try:
		local_BaseLineAlignment = retVal.BaseLineAlignment
	except:
		local_BaseLineAlignment = None
	try:
		local_AutoAdjustRightIndent = retVal.AutoAdjustRightIndent
	except:
		local_AutoAdjustRightIndent = None
	try:
		local_DisableLineHeightGrid = retVal.DisableLineHeightGrid
	except:
		local_DisableLineHeightGrid = None
	try:
		local_OutlineLevel = retVal.OutlineLevel
	except:
		local_OutlineLevel = None
	try:
		local_CharacterUnitRightIndent = retVal.CharacterUnitRightIndent
	except:
		local_CharacterUnitRightIndent = None
	try:
		local_CharacterUnitLeftIndent = retVal.CharacterUnitLeftIndent
	except:
		local_CharacterUnitLeftIndent = None
	try:
		local_CharacterUnitFirstLineIndent = retVal.CharacterUnitFirstLineIndent
	except:
		local_CharacterUnitFirstLineIndent = None
	try:
		local_LineUnitBefore = retVal.LineUnitBefore
	except:
		local_LineUnitBefore = None
	try:
		local_LineUnitAfter = retVal.LineUnitAfter
	except:
		local_LineUnitAfter = None
	try:
		local_ReadingOrder = retVal.ReadingOrder
	except:
		local_ReadingOrder = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_SpaceBeforeAuto = retVal.SpaceBeforeAuto
	except:
		local_SpaceBeforeAuto = None
	try:
		local_SpaceAfterAuto = retVal.SpaceAfterAuto
	except:
		local_SpaceAfterAuto = None
	try:
		local_IsStyleSeparator = retVal.IsStyleSeparator
	except:
		local_IsStyleSeparator = None
	try:
		local_MirrorIndents = retVal.MirrorIndents
	except:
		local_MirrorIndents = None
	try:
		local_TextboxTightWrap = retVal.TextboxTightWrap
	except:
		local_TextboxTightWrap = None
	try:
		local_ListNumberOriginal = retVal.ListNumberOriginal
	except:
		local_ListNumberOriginal = None
	try:
		local_ParaID = retVal.ParaID
	except:
		local_ParaID = None
	try:
		local_TextID = retVal.TextID
	except:
		local_TextID = None
	try:
		local_CollapsedState = retVal.CollapsedState
	except:
		local_CollapsedState = None
	try:
		local_CollapseHeadingByDefault = retVal.CollapseHeadingByDefault
	except:
		local_CollapseHeadingByDefault = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Paragraph", "Style": local_Style, "Alignment": local_Alignment, "KeepTogether": local_KeepTogether, "KeepWithNext": local_KeepWithNext, "PageBreakBefore": local_PageBreakBefore, "NoLineNumber": local_NoLineNumber, "RightIndent": local_RightIndent, "LeftIndent": local_LeftIndent, "FirstLineIndent": local_FirstLineIndent, "LineSpacing": local_LineSpacing, "LineSpacingRule": local_LineSpacingRule, "SpaceBefore": local_SpaceBefore, "SpaceAfter": local_SpaceAfter, "Hyphenation": local_Hyphenation, "WidowControl": local_WidowControl, "FarEastLineBreakControl": local_FarEastLineBreakControl, "WordWrap": local_WordWrap, "HangingPunctuation": local_HangingPunctuation, "HalfWidthPunctuationOnTopOfLine": local_HalfWidthPunctuationOnTopOfLine, "AddSpaceBetweenFarEastAndAlpha": local_AddSpaceBetweenFarEastAndAlpha, "AddSpaceBetweenFarEastAndDigit": local_AddSpaceBetweenFarEastAndDigit, "BaseLineAlignment": local_BaseLineAlignment, "AutoAdjustRightIndent": local_AutoAdjustRightIndent, "DisableLineHeightGrid": local_DisableLineHeightGrid, "OutlineLevel": local_OutlineLevel, "CharacterUnitRightIndent": local_CharacterUnitRightIndent, "CharacterUnitLeftIndent": local_CharacterUnitLeftIndent, "CharacterUnitFirstLineIndent": local_CharacterUnitFirstLineIndent, "LineUnitBefore": local_LineUnitBefore, "LineUnitAfter": local_LineUnitAfter, "ReadingOrder": local_ReadingOrder, "ID": local_ID, "SpaceBeforeAuto": local_SpaceBeforeAuto, "SpaceAfterAuto": local_SpaceAfterAuto, "IsStyleSeparator": local_IsStyleSeparator, "MirrorIndents": local_MirrorIndents, "TextboxTightWrap": local_TextboxTightWrap, "ListNumberOriginal": local_ListNumberOriginal, "ParaID": local_ParaID, "TextID": local_TextID, "CollapsedState": local_CollapsedState, "CollapseHeadingByDefault": local_CollapseHeadingByDefault, }


# Tool: 311
@mcp.tool()
async def word_Paragraph_OutlinePromote(this_Paragraph_wordObjId: str):
	"""This tool calls the OutlinePromote methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.OutlinePromote()


# Tool: 312
@mcp.tool()
async def word_Paragraph_OutlineDemote(this_Paragraph_wordObjId: str):
	"""This tool calls the OutlineDemote methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.OutlineDemote()


# Tool: 313
@mcp.tool()
async def word_Paragraph_OutlineDemoteToBody(this_Paragraph_wordObjId: str):
	"""This tool calls the OutlineDemoteToBody methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.OutlineDemoteToBody()


# Tool: 314
@mcp.tool()
async def word_Paragraph_Indent(this_Paragraph_wordObjId: str):
	"""This tool calls the Indent methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.Indent()


# Tool: 315
@mcp.tool()
async def word_Paragraph_Outdent(this_Paragraph_wordObjId: str):
	"""This tool calls the Outdent methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.Outdent()


# Tool: 316
@mcp.tool()
async def word_Paragraph_SelectNumber(this_Paragraph_wordObjId: str):
	"""This tool calls the SelectNumber methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.SelectNumber()


# Tool: 317
@mcp.tool()
async def word_Paragraph_ListAdvanceTo(this_Paragraph_wordObjId: str, Level1: int, Level2: int, Level3: int, Level4: int, Level5: int, Level6: int, Level7: int, Level8: int, Level9: int):
	"""This tool calls the ListAdvanceTo methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
	
	Parameters:
		Level1: the Level1 as int
		Level2: the Level2 as int
		Level3: the Level3 as int
		Level4: the Level4 as int
		Level5: the Level5 as int
		Level6: the Level6 as int
		Level7: the Level7 as int
		Level8: the Level8 as int
		Level9: the Level9 as int
	"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.ListAdvanceTo(Level1, Level2, Level3, Level4, Level5, Level6, Level7, Level8, Level9)


# Tool: 318
@mcp.tool()
async def word_Paragraph_ResetAdvanceTo(this_Paragraph_wordObjId: str):
	"""This tool calls the ResetAdvanceTo methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.ResetAdvanceTo()


# Tool: 319
@mcp.tool()
async def word_Paragraph_SeparateList(this_Paragraph_wordObjId: str):
	"""This tool calls the SeparateList methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.SeparateList()


# Tool: 320
@mcp.tool()
async def word_Paragraph_JoinList(this_Paragraph_wordObjId: str):
	"""This tool calls the JoinList methodon an Paragraph object. Pass the __WordObjectId of Paragraph of the object you want to call the method on as the first parameter
"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	this_Paragraph.JoinList()


# Tool: 321
@mcp.tool()
async def word_Paragraph_get_Property(this_Paragraph_wordObjId: str, propertyName: str):
	"""Gets properties of Paragraph
	
	propertyName: Name of the property. Can be one of ...
		Range, Format, TabStops, Borders, DropCap, Style, Alignment, KeepTogether, KeepWithNext, PageBreakBefore, NoLineNumber, RightIndent, LeftIndent, FirstLineIndent, LineSpacing, LineSpacingRule, SpaceBefore, SpaceAfter, Hyphenation, WidowControl, Shading, FarEastLineBreakControl, WordWrap, HangingPunctuation, HalfWidthPunctuationOnTopOfLine, AddSpaceBetweenFarEastAndAlpha, AddSpaceBetweenFarEastAndDigit, BaseLineAlignment, AutoAdjustRightIndent, DisableLineHeightGrid, OutlineLevel, CharacterUnitRightIndent, CharacterUnitLeftIndent, CharacterUnitFirstLineIndent, LineUnitBefore, LineUnitAfter, ReadingOrder, ID, SpaceBeforeAuto, SpaceAfterAuto, IsStyleSeparator, MirrorIndents, TextboxTightWrap, ParaID, TextID, CollapsedState, CollapseHeadingByDefault
	"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	
	EnsureWord()
	if (propertyName == "Range"):
		retVal = this_Paragraph.Range
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_Start = retVal.Start
		except:
			local_Start = None
		try:
			local_End = retVal.End
		except:
			local_End = None
		try:
			local_StoryType = retVal.StoryType
		except:
			local_StoryType = None
		try:
			local_Bold = retVal.Bold
		except:
			local_Bold = None
		try:
			local_Italic = retVal.Italic
		except:
			local_Italic = None
		try:
			local_Underline = retVal.Underline
		except:
			local_Underline = None
		try:
			local_EmphasisMark = retVal.EmphasisMark
		except:
			local_EmphasisMark = None
		try:
			local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
		except:
			local_DisableCharacterSpaceGrid = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_StoryLength = retVal.StoryLength
		except:
			local_StoryLength = None
		try:
			local_LanguageID = retVal.LanguageID
		except:
			local_LanguageID = None
		try:
			local_GrammarChecked = retVal.GrammarChecked
		except:
			local_GrammarChecked = None
		try:
			local_SpellingChecked = retVal.SpellingChecked
		except:
			local_SpellingChecked = None
		try:
			local_HighlightColorIndex = retVal.HighlightColorIndex
		except:
			local_HighlightColorIndex = None
		try:
			local_CanEdit = retVal.CanEdit
		except:
			local_CanEdit = None
		try:
			local_CanPaste = retVal.CanPaste
		except:
			local_CanPaste = None
		try:
			local_IsEndOfRowMark = retVal.IsEndOfRowMark
		except:
			local_IsEndOfRowMark = None
		try:
			local_BookmarkID = retVal.BookmarkID
		except:
			local_BookmarkID = None
		try:
			local_PreviousBookmarkID = retVal.PreviousBookmarkID
		except:
			local_PreviousBookmarkID = None
		try:
			local_Case = retVal.Case
		except:
			local_Case = None
		try:
			local_Information = retVal.Information
		except:
			local_Information = None
		try:
			local_Orientation = retVal.Orientation
		except:
			local_Orientation = None
		try:
			local_LanguageIDFarEast = retVal.LanguageIDFarEast
		except:
			local_LanguageIDFarEast = None
		try:
			local_LanguageIDOther = retVal.LanguageIDOther
		except:
			local_LanguageIDOther = None
		try:
			local_LanguageDetected = retVal.LanguageDetected
		except:
			local_LanguageDetected = None
		try:
			local_FitTextWidth = retVal.FitTextWidth
		except:
			local_FitTextWidth = None
		try:
			local_HorizontalInVertical = retVal.HorizontalInVertical
		except:
			local_HorizontalInVertical = None
		try:
			local_TwoLinesInOne = retVal.TwoLinesInOne
		except:
			local_TwoLinesInOne = None
		try:
			local_CombineCharacters = retVal.CombineCharacters
		except:
			local_CombineCharacters = None
		try:
			local_NoProofing = retVal.NoProofing
		except:
			local_NoProofing = None
		try:
			local_CharacterWidth = retVal.CharacterWidth
		except:
			local_CharacterWidth = None
		try:
			local_Kana = retVal.Kana
		except:
			local_Kana = None
		try:
			local_BoldBi = retVal.BoldBi
		except:
			local_BoldBi = None
		try:
			local_ItalicBi = retVal.ItalicBi
		except:
			local_ItalicBi = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_ShowAll = retVal.ShowAll
		except:
			local_ShowAll = None
		try:
			local_CharacterStyle = retVal.CharacterStyle
		except:
			local_CharacterStyle = None
		try:
			local_ParagraphStyle = retVal.ParagraphStyle
		except:
			local_ParagraphStyle = None
		try:
			local_ListStyle = retVal.ListStyle
		except:
			local_ListStyle = None
		try:
			local_TableStyle = retVal.TableStyle
		except:
			local_TableStyle = None
		try:
			local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
		except:
			local_TextVisibleOnScreen = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }
	if (propertyName == "Format"):
		retVal = this_Paragraph.Format
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ParagraphFormat"}
	if (propertyName == "TabStops"):
		retVal = this_Paragraph.TabStops
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "TabStops", "Count": local_Count, }
	if (propertyName == "Borders"):
		retVal = this_Paragraph.Borders
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Enable = retVal.Enable
		except:
			local_Enable = None
		try:
			local_DistanceFromTop = retVal.DistanceFromTop
		except:
			local_DistanceFromTop = None
		try:
			local_Shadow = retVal.Shadow
		except:
			local_Shadow = None
		try:
			local_InsideLineStyle = retVal.InsideLineStyle
		except:
			local_InsideLineStyle = None
		try:
			local_OutsideLineStyle = retVal.OutsideLineStyle
		except:
			local_OutsideLineStyle = None
		try:
			local_InsideLineWidth = retVal.InsideLineWidth
		except:
			local_InsideLineWidth = None
		try:
			local_OutsideLineWidth = retVal.OutsideLineWidth
		except:
			local_OutsideLineWidth = None
		try:
			local_InsideColorIndex = retVal.InsideColorIndex
		except:
			local_InsideColorIndex = None
		try:
			local_OutsideColorIndex = retVal.OutsideColorIndex
		except:
			local_OutsideColorIndex = None
		try:
			local_DistanceFromLeft = retVal.DistanceFromLeft
		except:
			local_DistanceFromLeft = None
		try:
			local_DistanceFromBottom = retVal.DistanceFromBottom
		except:
			local_DistanceFromBottom = None
		try:
			local_DistanceFromRight = retVal.DistanceFromRight
		except:
			local_DistanceFromRight = None
		try:
			local_AlwaysInFront = retVal.AlwaysInFront
		except:
			local_AlwaysInFront = None
		try:
			local_SurroundHeader = retVal.SurroundHeader
		except:
			local_SurroundHeader = None
		try:
			local_SurroundFooter = retVal.SurroundFooter
		except:
			local_SurroundFooter = None
		try:
			local_JoinBorders = retVal.JoinBorders
		except:
			local_JoinBorders = None
		try:
			local_HasHorizontal = retVal.HasHorizontal
		except:
			local_HasHorizontal = None
		try:
			local_HasVertical = retVal.HasVertical
		except:
			local_HasVertical = None
		try:
			local_DistanceFrom = retVal.DistanceFrom
		except:
			local_DistanceFrom = None
		try:
			local_EnableFirstPageInSection = retVal.EnableFirstPageInSection
		except:
			local_EnableFirstPageInSection = None
		try:
			local_EnableOtherPagesInSection = retVal.EnableOtherPagesInSection
		except:
			local_EnableOtherPagesInSection = None
		try:
			local_InsideColor = retVal.InsideColor
		except:
			local_InsideColor = None
		try:
			local_OutsideColor = retVal.OutsideColor
		except:
			local_OutsideColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Borders", "Count": local_Count, "Enable": local_Enable, "DistanceFromTop": local_DistanceFromTop, "Shadow": local_Shadow, "InsideLineStyle": local_InsideLineStyle, "OutsideLineStyle": local_OutsideLineStyle, "InsideLineWidth": local_InsideLineWidth, "OutsideLineWidth": local_OutsideLineWidth, "InsideColorIndex": local_InsideColorIndex, "OutsideColorIndex": local_OutsideColorIndex, "DistanceFromLeft": local_DistanceFromLeft, "DistanceFromBottom": local_DistanceFromBottom, "DistanceFromRight": local_DistanceFromRight, "AlwaysInFront": local_AlwaysInFront, "SurroundHeader": local_SurroundHeader, "SurroundFooter": local_SurroundFooter, "JoinBorders": local_JoinBorders, "HasHorizontal": local_HasHorizontal, "HasVertical": local_HasVertical, "DistanceFrom": local_DistanceFrom, "EnableFirstPageInSection": local_EnableFirstPageInSection, "EnableOtherPagesInSection": local_EnableOtherPagesInSection, "InsideColor": local_InsideColor, "OutsideColor": local_OutsideColor, }
	if (propertyName == "DropCap"):
		retVal = this_Paragraph.DropCap
		try:
			local_Position = retVal.Position
		except:
			local_Position = None
		try:
			local_FontName = retVal.FontName
		except:
			local_FontName = None
		try:
			local_LinesToDrop = retVal.LinesToDrop
		except:
			local_LinesToDrop = None
		try:
			local_DistanceFromText = retVal.DistanceFromText
		except:
			local_DistanceFromText = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "DropCap", "Position": local_Position, "FontName": local_FontName, "LinesToDrop": local_LinesToDrop, "DistanceFromText": local_DistanceFromText, }
	if (propertyName == "Style"):
		retVal = this_Paragraph.Style
		return retVal
	if (propertyName == "Alignment"):
		retVal = this_Paragraph.Alignment
		return retVal
	if (propertyName == "KeepTogether"):
		retVal = this_Paragraph.KeepTogether
		return retVal
	if (propertyName == "KeepWithNext"):
		retVal = this_Paragraph.KeepWithNext
		return retVal
	if (propertyName == "PageBreakBefore"):
		retVal = this_Paragraph.PageBreakBefore
		return retVal
	if (propertyName == "NoLineNumber"):
		retVal = this_Paragraph.NoLineNumber
		return retVal
	if (propertyName == "RightIndent"):
		retVal = this_Paragraph.RightIndent
		return retVal
	if (propertyName == "LeftIndent"):
		retVal = this_Paragraph.LeftIndent
		return retVal
	if (propertyName == "FirstLineIndent"):
		retVal = this_Paragraph.FirstLineIndent
		return retVal
	if (propertyName == "LineSpacing"):
		retVal = this_Paragraph.LineSpacing
		return retVal
	if (propertyName == "LineSpacingRule"):
		retVal = this_Paragraph.LineSpacingRule
		return retVal
	if (propertyName == "SpaceBefore"):
		retVal = this_Paragraph.SpaceBefore
		return retVal
	if (propertyName == "SpaceAfter"):
		retVal = this_Paragraph.SpaceAfter
		return retVal
	if (propertyName == "Hyphenation"):
		retVal = this_Paragraph.Hyphenation
		return retVal
	if (propertyName == "WidowControl"):
		retVal = this_Paragraph.WidowControl
		return retVal
	if (propertyName == "Shading"):
		retVal = this_Paragraph.Shading
		try:
			local_ForegroundPatternColorIndex = retVal.ForegroundPatternColorIndex
		except:
			local_ForegroundPatternColorIndex = None
		try:
			local_BackgroundPatternColorIndex = retVal.BackgroundPatternColorIndex
		except:
			local_BackgroundPatternColorIndex = None
		try:
			local_Texture = retVal.Texture
		except:
			local_Texture = None
		try:
			local_ForegroundPatternColor = retVal.ForegroundPatternColor
		except:
			local_ForegroundPatternColor = None
		try:
			local_BackgroundPatternColor = retVal.BackgroundPatternColor
		except:
			local_BackgroundPatternColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Shading", "ForegroundPatternColorIndex": local_ForegroundPatternColorIndex, "BackgroundPatternColorIndex": local_BackgroundPatternColorIndex, "Texture": local_Texture, "ForegroundPatternColor": local_ForegroundPatternColor, "BackgroundPatternColor": local_BackgroundPatternColor, }
	if (propertyName == "FarEastLineBreakControl"):
		retVal = this_Paragraph.FarEastLineBreakControl
		return retVal
	if (propertyName == "WordWrap"):
		retVal = this_Paragraph.WordWrap
		return retVal
	if (propertyName == "HangingPunctuation"):
		retVal = this_Paragraph.HangingPunctuation
		return retVal
	if (propertyName == "HalfWidthPunctuationOnTopOfLine"):
		retVal = this_Paragraph.HalfWidthPunctuationOnTopOfLine
		return retVal
	if (propertyName == "AddSpaceBetweenFarEastAndAlpha"):
		retVal = this_Paragraph.AddSpaceBetweenFarEastAndAlpha
		return retVal
	if (propertyName == "AddSpaceBetweenFarEastAndDigit"):
		retVal = this_Paragraph.AddSpaceBetweenFarEastAndDigit
		return retVal
	if (propertyName == "BaseLineAlignment"):
		retVal = this_Paragraph.BaseLineAlignment
		return retVal
	if (propertyName == "AutoAdjustRightIndent"):
		retVal = this_Paragraph.AutoAdjustRightIndent
		return retVal
	if (propertyName == "DisableLineHeightGrid"):
		retVal = this_Paragraph.DisableLineHeightGrid
		return retVal
	if (propertyName == "OutlineLevel"):
		retVal = this_Paragraph.OutlineLevel
		return retVal
	if (propertyName == "CharacterUnitRightIndent"):
		retVal = this_Paragraph.CharacterUnitRightIndent
		return retVal
	if (propertyName == "CharacterUnitLeftIndent"):
		retVal = this_Paragraph.CharacterUnitLeftIndent
		return retVal
	if (propertyName == "CharacterUnitFirstLineIndent"):
		retVal = this_Paragraph.CharacterUnitFirstLineIndent
		return retVal
	if (propertyName == "LineUnitBefore"):
		retVal = this_Paragraph.LineUnitBefore
		return retVal
	if (propertyName == "LineUnitAfter"):
		retVal = this_Paragraph.LineUnitAfter
		return retVal
	if (propertyName == "ReadingOrder"):
		retVal = this_Paragraph.ReadingOrder
		return retVal
	if (propertyName == "ID"):
		retVal = this_Paragraph.ID
		return retVal
	if (propertyName == "SpaceBeforeAuto"):
		retVal = this_Paragraph.SpaceBeforeAuto
		return retVal
	if (propertyName == "SpaceAfterAuto"):
		retVal = this_Paragraph.SpaceAfterAuto
		return retVal
	if (propertyName == "IsStyleSeparator"):
		retVal = this_Paragraph.IsStyleSeparator
		return retVal
	if (propertyName == "MirrorIndents"):
		retVal = this_Paragraph.MirrorIndents
		return retVal
	if (propertyName == "TextboxTightWrap"):
		retVal = this_Paragraph.TextboxTightWrap
		return retVal
	if (propertyName == "ParaID"):
		retVal = this_Paragraph.ParaID
		return retVal
	if (propertyName == "TextID"):
		retVal = this_Paragraph.TextID
		return retVal
	if (propertyName == "CollapsedState"):
		retVal = this_Paragraph.CollapsedState
		return retVal
	if (propertyName == "CollapseHeadingByDefault"):
		retVal = this_Paragraph.CollapseHeadingByDefault
		return retVal


# Tool: 322
@mcp.tool()
async def word_Paragraph_set_Property(this_Paragraph_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Paragraph
	
	propertyName: Name of the property. Can be one of ...
		Format, TabStops, Borders, Style, Alignment, KeepTogether, KeepWithNext, PageBreakBefore, NoLineNumber, RightIndent, LeftIndent, FirstLineIndent, LineSpacing, LineSpacingRule, SpaceBefore, SpaceAfter, Hyphenation, WidowControl, FarEastLineBreakControl, WordWrap, HangingPunctuation, HalfWidthPunctuationOnTopOfLine, AddSpaceBetweenFarEastAndAlpha, AddSpaceBetweenFarEastAndDigit, BaseLineAlignment, AutoAdjustRightIndent, DisableLineHeightGrid, OutlineLevel, CharacterUnitRightIndent, CharacterUnitLeftIndent, CharacterUnitFirstLineIndent, LineUnitBefore, LineUnitAfter, ReadingOrder, ID, SpaceBeforeAuto, SpaceAfterAuto, MirrorIndents, TextboxTightWrap, CollapsedState, CollapseHeadingByDefault
	"""
	this_Paragraph = get_object(this_Paragraph_wordObjId)
	
	EnsureWord()
	if (propertyName == "Format"):
		this_Paragraph.Format = propertyValue
	if (propertyName == "TabStops"):
		this_Paragraph.TabStops = propertyValue
	if (propertyName == "Borders"):
		this_Paragraph.Borders = propertyValue
	if (propertyName == "Style"):
		this_Paragraph.Style = propertyValue
	if (propertyName == "Alignment"):
		this_Paragraph.Alignment = propertyValue
	if (propertyName == "KeepTogether"):
		this_Paragraph.KeepTogether = propertyValue
	if (propertyName == "KeepWithNext"):
		this_Paragraph.KeepWithNext = propertyValue
	if (propertyName == "PageBreakBefore"):
		this_Paragraph.PageBreakBefore = propertyValue
	if (propertyName == "NoLineNumber"):
		this_Paragraph.NoLineNumber = propertyValue
	if (propertyName == "RightIndent"):
		this_Paragraph.RightIndent = propertyValue
	if (propertyName == "LeftIndent"):
		this_Paragraph.LeftIndent = propertyValue
	if (propertyName == "FirstLineIndent"):
		this_Paragraph.FirstLineIndent = propertyValue
	if (propertyName == "LineSpacing"):
		this_Paragraph.LineSpacing = propertyValue
	if (propertyName == "LineSpacingRule"):
		this_Paragraph.LineSpacingRule = propertyValue
	if (propertyName == "SpaceBefore"):
		this_Paragraph.SpaceBefore = propertyValue
	if (propertyName == "SpaceAfter"):
		this_Paragraph.SpaceAfter = propertyValue
	if (propertyName == "Hyphenation"):
		this_Paragraph.Hyphenation = propertyValue
	if (propertyName == "WidowControl"):
		this_Paragraph.WidowControl = propertyValue
	if (propertyName == "FarEastLineBreakControl"):
		this_Paragraph.FarEastLineBreakControl = propertyValue
	if (propertyName == "WordWrap"):
		this_Paragraph.WordWrap = propertyValue
	if (propertyName == "HangingPunctuation"):
		this_Paragraph.HangingPunctuation = propertyValue
	if (propertyName == "HalfWidthPunctuationOnTopOfLine"):
		this_Paragraph.HalfWidthPunctuationOnTopOfLine = propertyValue
	if (propertyName == "AddSpaceBetweenFarEastAndAlpha"):
		this_Paragraph.AddSpaceBetweenFarEastAndAlpha = propertyValue
	if (propertyName == "AddSpaceBetweenFarEastAndDigit"):
		this_Paragraph.AddSpaceBetweenFarEastAndDigit = propertyValue
	if (propertyName == "BaseLineAlignment"):
		this_Paragraph.BaseLineAlignment = propertyValue
	if (propertyName == "AutoAdjustRightIndent"):
		this_Paragraph.AutoAdjustRightIndent = propertyValue
	if (propertyName == "DisableLineHeightGrid"):
		this_Paragraph.DisableLineHeightGrid = propertyValue
	if (propertyName == "OutlineLevel"):
		this_Paragraph.OutlineLevel = propertyValue
	if (propertyName == "CharacterUnitRightIndent"):
		this_Paragraph.CharacterUnitRightIndent = propertyValue
	if (propertyName == "CharacterUnitLeftIndent"):
		this_Paragraph.CharacterUnitLeftIndent = propertyValue
	if (propertyName == "CharacterUnitFirstLineIndent"):
		this_Paragraph.CharacterUnitFirstLineIndent = propertyValue
	if (propertyName == "LineUnitBefore"):
		this_Paragraph.LineUnitBefore = propertyValue
	if (propertyName == "LineUnitAfter"):
		this_Paragraph.LineUnitAfter = propertyValue
	if (propertyName == "ReadingOrder"):
		this_Paragraph.ReadingOrder = propertyValue
	if (propertyName == "ID"):
		this_Paragraph.ID = propertyValue
	if (propertyName == "SpaceBeforeAuto"):
		this_Paragraph.SpaceBeforeAuto = propertyValue
	if (propertyName == "SpaceAfterAuto"):
		this_Paragraph.SpaceAfterAuto = propertyValue
	if (propertyName == "MirrorIndents"):
		this_Paragraph.MirrorIndents = propertyValue
	if (propertyName == "TextboxTightWrap"):
		this_Paragraph.TextboxTightWrap = propertyValue
	if (propertyName == "CollapsedState"):
		this_Paragraph.CollapsedState = propertyValue
	if (propertyName == "CollapseHeadingByDefault"):
		this_Paragraph.CollapseHeadingByDefault = propertyValue


# Tool: 323
@mcp.tool()
async def word_Table_Select(this_Table_wordObjId: str):
	"""This tool calls the Select methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
"""
	this_Table = get_object(this_Table_wordObjId)
	this_Table.Select()


# Tool: 324
@mcp.tool()
async def word_Table_Delete(this_Table_wordObjId: str):
	"""This tool calls the Delete methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
"""
	this_Table = get_object(this_Table_wordObjId)
	this_Table.Delete()


# Tool: 325
@mcp.tool()
async def word_Table_SortOld(this_Table_wordObjId: str, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, CaseSensitive, LanguageID):
	"""This tool calls the SortOld methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
	
	Parameters:
		ExcludeHeader: the ExcludeHeader as VT_VARIANT
		FieldNumber: the FieldNumber as VT_VARIANT
		SortFieldType: the SortFieldType as VT_VARIANT
		SortOrder: the SortOrder as VT_VARIANT
		FieldNumber2: the FieldNumber2 as VT_VARIANT
		SortFieldType2: the SortFieldType2 as VT_VARIANT
		SortOrder2: the SortOrder2 as VT_VARIANT
		FieldNumber3: the FieldNumber3 as VT_VARIANT
		SortFieldType3: the SortFieldType3 as VT_VARIANT
		SortOrder3: the SortOrder3 as VT_VARIANT
		CaseSensitive: the CaseSensitive as VT_VARIANT
		LanguageID: the LanguageID as VT_VARIANT
	"""
	this_Table = get_object(this_Table_wordObjId)
	ExcludeHeader = tryParseString(ExcludeHeader)
	FieldNumber = tryParseString(FieldNumber)
	SortFieldType = tryParseString(SortFieldType)
	SortOrder = tryParseString(SortOrder)
	FieldNumber2 = tryParseString(FieldNumber2)
	SortFieldType2 = tryParseString(SortFieldType2)
	SortOrder2 = tryParseString(SortOrder2)
	FieldNumber3 = tryParseString(FieldNumber3)
	SortFieldType3 = tryParseString(SortFieldType3)
	SortOrder3 = tryParseString(SortOrder3)
	CaseSensitive = tryParseString(CaseSensitive)
	LanguageID = tryParseString(LanguageID)
	this_Table.SortOld(ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, CaseSensitive, LanguageID)


# Tool: 326
@mcp.tool()
async def word_Table_SortAscending(this_Table_wordObjId: str):
	"""This tool calls the SortAscending methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
"""
	this_Table = get_object(this_Table_wordObjId)
	this_Table.SortAscending()


# Tool: 327
@mcp.tool()
async def word_Table_SortDescending(this_Table_wordObjId: str):
	"""This tool calls the SortDescending methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
"""
	this_Table = get_object(this_Table_wordObjId)
	this_Table.SortDescending()


# Tool: 328
@mcp.tool()
async def word_Table_AutoFormat(this_Table_wordObjId: str, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit):
	"""This tool calls the AutoFormat methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
	
	Parameters:
		Format: the Format as VT_VARIANT
		ApplyBorders: the ApplyBorders as VT_VARIANT
		ApplyShading: the ApplyShading as VT_VARIANT
		ApplyFont: the ApplyFont as VT_VARIANT
		ApplyColor: the ApplyColor as VT_VARIANT
		ApplyHeadingRows: the ApplyHeadingRows as VT_VARIANT
		ApplyLastRow: the ApplyLastRow as VT_VARIANT
		ApplyFirstColumn: the ApplyFirstColumn as VT_VARIANT
		ApplyLastColumn: the ApplyLastColumn as VT_VARIANT
		AutoFit: the AutoFit as VT_VARIANT
	"""
	this_Table = get_object(this_Table_wordObjId)
	Format = tryParseString(Format)
	ApplyBorders = tryParseString(ApplyBorders)
	ApplyShading = tryParseString(ApplyShading)
	ApplyFont = tryParseString(ApplyFont)
	ApplyColor = tryParseString(ApplyColor)
	ApplyHeadingRows = tryParseString(ApplyHeadingRows)
	ApplyLastRow = tryParseString(ApplyLastRow)
	ApplyFirstColumn = tryParseString(ApplyFirstColumn)
	ApplyLastColumn = tryParseString(ApplyLastColumn)
	AutoFit = tryParseString(AutoFit)
	this_Table.AutoFormat(Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit)


# Tool: 329
@mcp.tool()
async def word_Table_UpdateAutoFormat(this_Table_wordObjId: str):
	"""This tool calls the UpdateAutoFormat methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
"""
	this_Table = get_object(this_Table_wordObjId)
	this_Table.UpdateAutoFormat()


# Tool: 330
@mcp.tool()
async def word_Table_ConvertToTextOld(this_Table_wordObjId: str, Separator):
	"""This tool calls the ConvertToTextOld methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
	
	Parameters:
		Separator: the Separator as VT_VARIANT
	"""
	this_Table = get_object(this_Table_wordObjId)
	Separator = tryParseString(Separator)
	retVal = this_Table.ConvertToTextOld(Separator)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 331
@mcp.tool()
async def word_Table_Cell(this_Table_wordObjId: str, Row: int, Column: int):
	"""This tool calls the Cell methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
	
	Parameters:
		Row: the Row as int
		Column: the Column as int
	"""
	this_Table = get_object(this_Table_wordObjId)
	retVal = this_Table.Cell(Row, Column)
	try:
		local_RowIndex = retVal.RowIndex
	except:
		local_RowIndex = None
	try:
		local_ColumnIndex = retVal.ColumnIndex
	except:
		local_ColumnIndex = None
	try:
		local_Width = retVal.Width
	except:
		local_Width = None
	try:
		local_Height = retVal.Height
	except:
		local_Height = None
	try:
		local_HeightRule = retVal.HeightRule
	except:
		local_HeightRule = None
	try:
		local_VerticalAlignment = retVal.VerticalAlignment
	except:
		local_VerticalAlignment = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_WordWrap = retVal.WordWrap
	except:
		local_WordWrap = None
	try:
		local_PreferredWidth = retVal.PreferredWidth
	except:
		local_PreferredWidth = None
	try:
		local_FitText = retVal.FitText
	except:
		local_FitText = None
	try:
		local_TopPadding = retVal.TopPadding
	except:
		local_TopPadding = None
	try:
		local_BottomPadding = retVal.BottomPadding
	except:
		local_BottomPadding = None
	try:
		local_LeftPadding = retVal.LeftPadding
	except:
		local_LeftPadding = None
	try:
		local_RightPadding = retVal.RightPadding
	except:
		local_RightPadding = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_PreferredWidthType = retVal.PreferredWidthType
	except:
		local_PreferredWidthType = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Cell", "RowIndex": local_RowIndex, "ColumnIndex": local_ColumnIndex, "Width": local_Width, "Height": local_Height, "HeightRule": local_HeightRule, "VerticalAlignment": local_VerticalAlignment, "NestingLevel": local_NestingLevel, "WordWrap": local_WordWrap, "PreferredWidth": local_PreferredWidth, "FitText": local_FitText, "TopPadding": local_TopPadding, "BottomPadding": local_BottomPadding, "LeftPadding": local_LeftPadding, "RightPadding": local_RightPadding, "ID": local_ID, "PreferredWidthType": local_PreferredWidthType, }


# Tool: 332
@mcp.tool()
async def word_Table_Split(this_Table_wordObjId: str, BeforeRow):
	"""This tool calls the Split methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
	
	Parameters:
		BeforeRow: the BeforeRow as VT_VARIANT
	"""
	this_Table = get_object(this_Table_wordObjId)
	BeforeRow = tryParseString(BeforeRow)
	retVal = this_Table.Split(BeforeRow)
	try:
		local_Uniform = retVal.Uniform
	except:
		local_Uniform = None
	try:
		local_AutoFormatType = retVal.AutoFormatType
	except:
		local_AutoFormatType = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_AllowPageBreaks = retVal.AllowPageBreaks
	except:
		local_AllowPageBreaks = None
	try:
		local_AllowAutoFit = retVal.AllowAutoFit
	except:
		local_AllowAutoFit = None
	try:
		local_PreferredWidth = retVal.PreferredWidth
	except:
		local_PreferredWidth = None
	try:
		local_PreferredWidthType = retVal.PreferredWidthType
	except:
		local_PreferredWidthType = None
	try:
		local_TopPadding = retVal.TopPadding
	except:
		local_TopPadding = None
	try:
		local_BottomPadding = retVal.BottomPadding
	except:
		local_BottomPadding = None
	try:
		local_LeftPadding = retVal.LeftPadding
	except:
		local_LeftPadding = None
	try:
		local_RightPadding = retVal.RightPadding
	except:
		local_RightPadding = None
	try:
		local_Spacing = retVal.Spacing
	except:
		local_Spacing = None
	try:
		local_TableDirection = retVal.TableDirection
	except:
		local_TableDirection = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_ApplyStyleHeadingRows = retVal.ApplyStyleHeadingRows
	except:
		local_ApplyStyleHeadingRows = None
	try:
		local_ApplyStyleLastRow = retVal.ApplyStyleLastRow
	except:
		local_ApplyStyleLastRow = None
	try:
		local_ApplyStyleFirstColumn = retVal.ApplyStyleFirstColumn
	except:
		local_ApplyStyleFirstColumn = None
	try:
		local_ApplyStyleLastColumn = retVal.ApplyStyleLastColumn
	except:
		local_ApplyStyleLastColumn = None
	try:
		local_ApplyStyleRowBands = retVal.ApplyStyleRowBands
	except:
		local_ApplyStyleRowBands = None
	try:
		local_ApplyStyleColumnBands = retVal.ApplyStyleColumnBands
	except:
		local_ApplyStyleColumnBands = None
	try:
		local_Title = retVal.Title
	except:
		local_Title = None
	try:
		local_Descr = retVal.Descr
	except:
		local_Descr = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Table", "Uniform": local_Uniform, "AutoFormatType": local_AutoFormatType, "NestingLevel": local_NestingLevel, "AllowPageBreaks": local_AllowPageBreaks, "AllowAutoFit": local_AllowAutoFit, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, "TopPadding": local_TopPadding, "BottomPadding": local_BottomPadding, "LeftPadding": local_LeftPadding, "RightPadding": local_RightPadding, "Spacing": local_Spacing, "TableDirection": local_TableDirection, "ID": local_ID, "Style": local_Style, "ApplyStyleHeadingRows": local_ApplyStyleHeadingRows, "ApplyStyleLastRow": local_ApplyStyleLastRow, "ApplyStyleFirstColumn": local_ApplyStyleFirstColumn, "ApplyStyleLastColumn": local_ApplyStyleLastColumn, "ApplyStyleRowBands": local_ApplyStyleRowBands, "ApplyStyleColumnBands": local_ApplyStyleColumnBands, "Title": local_Title, "Descr": local_Descr, }


# Tool: 333
@mcp.tool()
async def word_Table_ConvertToText(this_Table_wordObjId: str, Separator, NestedTables):
	"""This tool calls the ConvertToText methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
	
	Parameters:
		Separator: the Separator as VT_VARIANT
		NestedTables: the NestedTables as VT_VARIANT
	"""
	this_Table = get_object(this_Table_wordObjId)
	Separator = tryParseString(Separator)
	NestedTables = tryParseString(NestedTables)
	retVal = this_Table.ConvertToText(Separator, NestedTables)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 334
@mcp.tool()
async def word_Table_AutoFitBehavior(this_Table_wordObjId: str, Behavior: int):
	"""This tool calls the AutoFitBehavior methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
	
	Parameters:
		Behavior: the Behavior as WdAutoFitBehavior
	"""
	this_Table = get_object(this_Table_wordObjId)
	this_Table.AutoFitBehavior(Behavior)


# Tool: 335
@mcp.tool()
async def word_Table_Sort(this_Table_wordObjId: str, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID):
	"""This tool calls the Sort methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
	
	Parameters:
		ExcludeHeader: the ExcludeHeader as VT_VARIANT
		FieldNumber: the FieldNumber as VT_VARIANT
		SortFieldType: the SortFieldType as VT_VARIANT
		SortOrder: the SortOrder as VT_VARIANT
		FieldNumber2: the FieldNumber2 as VT_VARIANT
		SortFieldType2: the SortFieldType2 as VT_VARIANT
		SortOrder2: the SortOrder2 as VT_VARIANT
		FieldNumber3: the FieldNumber3 as VT_VARIANT
		SortFieldType3: the SortFieldType3 as VT_VARIANT
		SortOrder3: the SortOrder3 as VT_VARIANT
		CaseSensitive: the CaseSensitive as VT_VARIANT
		BidiSort: the BidiSort as VT_VARIANT
		IgnoreThe: the IgnoreThe as VT_VARIANT
		IgnoreKashida: the IgnoreKashida as VT_VARIANT
		IgnoreDiacritics: the IgnoreDiacritics as VT_VARIANT
		IgnoreHe: the IgnoreHe as VT_VARIANT
		LanguageID: the LanguageID as VT_VARIANT
	"""
	this_Table = get_object(this_Table_wordObjId)
	ExcludeHeader = tryParseString(ExcludeHeader)
	FieldNumber = tryParseString(FieldNumber)
	SortFieldType = tryParseString(SortFieldType)
	SortOrder = tryParseString(SortOrder)
	FieldNumber2 = tryParseString(FieldNumber2)
	SortFieldType2 = tryParseString(SortFieldType2)
	SortOrder2 = tryParseString(SortOrder2)
	FieldNumber3 = tryParseString(FieldNumber3)
	SortFieldType3 = tryParseString(SortFieldType3)
	SortOrder3 = tryParseString(SortOrder3)
	CaseSensitive = tryParseString(CaseSensitive)
	BidiSort = tryParseString(BidiSort)
	IgnoreThe = tryParseString(IgnoreThe)
	IgnoreKashida = tryParseString(IgnoreKashida)
	IgnoreDiacritics = tryParseString(IgnoreDiacritics)
	IgnoreHe = tryParseString(IgnoreHe)
	LanguageID = tryParseString(LanguageID)
	this_Table.Sort(ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID)


# Tool: 336
@mcp.tool()
async def word_Table_ApplyStyleDirectFormatting(this_Table_wordObjId: str, StyleName: str):
	"""This tool calls the ApplyStyleDirectFormatting methodon an Table object. Pass the __WordObjectId of Table of the object you want to call the method on as the first parameter
	
	Parameters:
		StyleName: the StyleName as str
	"""
	this_Table = get_object(this_Table_wordObjId)
	this_Table.ApplyStyleDirectFormatting(StyleName)


# Tool: 337
@mcp.tool()
async def word_Table_get_Property(this_Table_wordObjId: str, propertyName: str):
	"""Gets properties of Table
	
	propertyName: Name of the property. Can be one of ...
		Range, Columns, Rows, Borders, Shading, Uniform, AutoFormatType, Tables, NestingLevel, AllowPageBreaks, AllowAutoFit, PreferredWidth, PreferredWidthType, TopPadding, BottomPadding, LeftPadding, RightPadding, Spacing, TableDirection, ID, Style, ApplyStyleHeadingRows, ApplyStyleLastRow, ApplyStyleFirstColumn, ApplyStyleLastColumn, ApplyStyleRowBands, ApplyStyleColumnBands, Title, Descr
	"""
	this_Table = get_object(this_Table_wordObjId)
	
	EnsureWord()
	if (propertyName == "Range"):
		retVal = this_Table.Range
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_Start = retVal.Start
		except:
			local_Start = None
		try:
			local_End = retVal.End
		except:
			local_End = None
		try:
			local_StoryType = retVal.StoryType
		except:
			local_StoryType = None
		try:
			local_Bold = retVal.Bold
		except:
			local_Bold = None
		try:
			local_Italic = retVal.Italic
		except:
			local_Italic = None
		try:
			local_Underline = retVal.Underline
		except:
			local_Underline = None
		try:
			local_EmphasisMark = retVal.EmphasisMark
		except:
			local_EmphasisMark = None
		try:
			local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
		except:
			local_DisableCharacterSpaceGrid = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_StoryLength = retVal.StoryLength
		except:
			local_StoryLength = None
		try:
			local_LanguageID = retVal.LanguageID
		except:
			local_LanguageID = None
		try:
			local_GrammarChecked = retVal.GrammarChecked
		except:
			local_GrammarChecked = None
		try:
			local_SpellingChecked = retVal.SpellingChecked
		except:
			local_SpellingChecked = None
		try:
			local_HighlightColorIndex = retVal.HighlightColorIndex
		except:
			local_HighlightColorIndex = None
		try:
			local_CanEdit = retVal.CanEdit
		except:
			local_CanEdit = None
		try:
			local_CanPaste = retVal.CanPaste
		except:
			local_CanPaste = None
		try:
			local_IsEndOfRowMark = retVal.IsEndOfRowMark
		except:
			local_IsEndOfRowMark = None
		try:
			local_BookmarkID = retVal.BookmarkID
		except:
			local_BookmarkID = None
		try:
			local_PreviousBookmarkID = retVal.PreviousBookmarkID
		except:
			local_PreviousBookmarkID = None
		try:
			local_Case = retVal.Case
		except:
			local_Case = None
		try:
			local_Information = retVal.Information
		except:
			local_Information = None
		try:
			local_Orientation = retVal.Orientation
		except:
			local_Orientation = None
		try:
			local_LanguageIDFarEast = retVal.LanguageIDFarEast
		except:
			local_LanguageIDFarEast = None
		try:
			local_LanguageIDOther = retVal.LanguageIDOther
		except:
			local_LanguageIDOther = None
		try:
			local_LanguageDetected = retVal.LanguageDetected
		except:
			local_LanguageDetected = None
		try:
			local_FitTextWidth = retVal.FitTextWidth
		except:
			local_FitTextWidth = None
		try:
			local_HorizontalInVertical = retVal.HorizontalInVertical
		except:
			local_HorizontalInVertical = None
		try:
			local_TwoLinesInOne = retVal.TwoLinesInOne
		except:
			local_TwoLinesInOne = None
		try:
			local_CombineCharacters = retVal.CombineCharacters
		except:
			local_CombineCharacters = None
		try:
			local_NoProofing = retVal.NoProofing
		except:
			local_NoProofing = None
		try:
			local_CharacterWidth = retVal.CharacterWidth
		except:
			local_CharacterWidth = None
		try:
			local_Kana = retVal.Kana
		except:
			local_Kana = None
		try:
			local_BoldBi = retVal.BoldBi
		except:
			local_BoldBi = None
		try:
			local_ItalicBi = retVal.ItalicBi
		except:
			local_ItalicBi = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_ShowAll = retVal.ShowAll
		except:
			local_ShowAll = None
		try:
			local_CharacterStyle = retVal.CharacterStyle
		except:
			local_CharacterStyle = None
		try:
			local_ParagraphStyle = retVal.ParagraphStyle
		except:
			local_ParagraphStyle = None
		try:
			local_ListStyle = retVal.ListStyle
		except:
			local_ListStyle = None
		try:
			local_TableStyle = retVal.TableStyle
		except:
			local_TableStyle = None
		try:
			local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
		except:
			local_TextVisibleOnScreen = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }
	if (propertyName == "Columns"):
		retVal = this_Table.Columns
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Columns", "Count": local_Count, "Width": local_Width, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "Rows"):
		retVal = this_Table.Rows
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_AllowBreakAcrossPages = retVal.AllowBreakAcrossPages
		except:
			local_AllowBreakAcrossPages = None
		try:
			local_Alignment = retVal.Alignment
		except:
			local_Alignment = None
		try:
			local_HeadingFormat = retVal.HeadingFormat
		except:
			local_HeadingFormat = None
		try:
			local_SpaceBetweenColumns = retVal.SpaceBetweenColumns
		except:
			local_SpaceBetweenColumns = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_LeftIndent = retVal.LeftIndent
		except:
			local_LeftIndent = None
		try:
			local_WrapAroundText = retVal.WrapAroundText
		except:
			local_WrapAroundText = None
		try:
			local_DistanceTop = retVal.DistanceTop
		except:
			local_DistanceTop = None
		try:
			local_DistanceBottom = retVal.DistanceBottom
		except:
			local_DistanceBottom = None
		try:
			local_DistanceLeft = retVal.DistanceLeft
		except:
			local_DistanceLeft = None
		try:
			local_DistanceRight = retVal.DistanceRight
		except:
			local_DistanceRight = None
		try:
			local_HorizontalPosition = retVal.HorizontalPosition
		except:
			local_HorizontalPosition = None
		try:
			local_VerticalPosition = retVal.VerticalPosition
		except:
			local_VerticalPosition = None
		try:
			local_RelativeHorizontalPosition = retVal.RelativeHorizontalPosition
		except:
			local_RelativeHorizontalPosition = None
		try:
			local_RelativeVerticalPosition = retVal.RelativeVerticalPosition
		except:
			local_RelativeVerticalPosition = None
		try:
			local_AllowOverlap = retVal.AllowOverlap
		except:
			local_AllowOverlap = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_TableDirection = retVal.TableDirection
		except:
			local_TableDirection = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Rows", "Count": local_Count, "AllowBreakAcrossPages": local_AllowBreakAcrossPages, "Alignment": local_Alignment, "HeadingFormat": local_HeadingFormat, "SpaceBetweenColumns": local_SpaceBetweenColumns, "Height": local_Height, "HeightRule": local_HeightRule, "LeftIndent": local_LeftIndent, "WrapAroundText": local_WrapAroundText, "DistanceTop": local_DistanceTop, "DistanceBottom": local_DistanceBottom, "DistanceLeft": local_DistanceLeft, "DistanceRight": local_DistanceRight, "HorizontalPosition": local_HorizontalPosition, "VerticalPosition": local_VerticalPosition, "RelativeHorizontalPosition": local_RelativeHorizontalPosition, "RelativeVerticalPosition": local_RelativeVerticalPosition, "AllowOverlap": local_AllowOverlap, "NestingLevel": local_NestingLevel, "TableDirection": local_TableDirection, }
	if (propertyName == "Borders"):
		retVal = this_Table.Borders
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Enable = retVal.Enable
		except:
			local_Enable = None
		try:
			local_DistanceFromTop = retVal.DistanceFromTop
		except:
			local_DistanceFromTop = None
		try:
			local_Shadow = retVal.Shadow
		except:
			local_Shadow = None
		try:
			local_InsideLineStyle = retVal.InsideLineStyle
		except:
			local_InsideLineStyle = None
		try:
			local_OutsideLineStyle = retVal.OutsideLineStyle
		except:
			local_OutsideLineStyle = None
		try:
			local_InsideLineWidth = retVal.InsideLineWidth
		except:
			local_InsideLineWidth = None
		try:
			local_OutsideLineWidth = retVal.OutsideLineWidth
		except:
			local_OutsideLineWidth = None
		try:
			local_InsideColorIndex = retVal.InsideColorIndex
		except:
			local_InsideColorIndex = None
		try:
			local_OutsideColorIndex = retVal.OutsideColorIndex
		except:
			local_OutsideColorIndex = None
		try:
			local_DistanceFromLeft = retVal.DistanceFromLeft
		except:
			local_DistanceFromLeft = None
		try:
			local_DistanceFromBottom = retVal.DistanceFromBottom
		except:
			local_DistanceFromBottom = None
		try:
			local_DistanceFromRight = retVal.DistanceFromRight
		except:
			local_DistanceFromRight = None
		try:
			local_AlwaysInFront = retVal.AlwaysInFront
		except:
			local_AlwaysInFront = None
		try:
			local_SurroundHeader = retVal.SurroundHeader
		except:
			local_SurroundHeader = None
		try:
			local_SurroundFooter = retVal.SurroundFooter
		except:
			local_SurroundFooter = None
		try:
			local_JoinBorders = retVal.JoinBorders
		except:
			local_JoinBorders = None
		try:
			local_HasHorizontal = retVal.HasHorizontal
		except:
			local_HasHorizontal = None
		try:
			local_HasVertical = retVal.HasVertical
		except:
			local_HasVertical = None
		try:
			local_DistanceFrom = retVal.DistanceFrom
		except:
			local_DistanceFrom = None
		try:
			local_EnableFirstPageInSection = retVal.EnableFirstPageInSection
		except:
			local_EnableFirstPageInSection = None
		try:
			local_EnableOtherPagesInSection = retVal.EnableOtherPagesInSection
		except:
			local_EnableOtherPagesInSection = None
		try:
			local_InsideColor = retVal.InsideColor
		except:
			local_InsideColor = None
		try:
			local_OutsideColor = retVal.OutsideColor
		except:
			local_OutsideColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Borders", "Count": local_Count, "Enable": local_Enable, "DistanceFromTop": local_DistanceFromTop, "Shadow": local_Shadow, "InsideLineStyle": local_InsideLineStyle, "OutsideLineStyle": local_OutsideLineStyle, "InsideLineWidth": local_InsideLineWidth, "OutsideLineWidth": local_OutsideLineWidth, "InsideColorIndex": local_InsideColorIndex, "OutsideColorIndex": local_OutsideColorIndex, "DistanceFromLeft": local_DistanceFromLeft, "DistanceFromBottom": local_DistanceFromBottom, "DistanceFromRight": local_DistanceFromRight, "AlwaysInFront": local_AlwaysInFront, "SurroundHeader": local_SurroundHeader, "SurroundFooter": local_SurroundFooter, "JoinBorders": local_JoinBorders, "HasHorizontal": local_HasHorizontal, "HasVertical": local_HasVertical, "DistanceFrom": local_DistanceFrom, "EnableFirstPageInSection": local_EnableFirstPageInSection, "EnableOtherPagesInSection": local_EnableOtherPagesInSection, "InsideColor": local_InsideColor, "OutsideColor": local_OutsideColor, }
	if (propertyName == "Shading"):
		retVal = this_Table.Shading
		try:
			local_ForegroundPatternColorIndex = retVal.ForegroundPatternColorIndex
		except:
			local_ForegroundPatternColorIndex = None
		try:
			local_BackgroundPatternColorIndex = retVal.BackgroundPatternColorIndex
		except:
			local_BackgroundPatternColorIndex = None
		try:
			local_Texture = retVal.Texture
		except:
			local_Texture = None
		try:
			local_ForegroundPatternColor = retVal.ForegroundPatternColor
		except:
			local_ForegroundPatternColor = None
		try:
			local_BackgroundPatternColor = retVal.BackgroundPatternColor
		except:
			local_BackgroundPatternColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Shading", "ForegroundPatternColorIndex": local_ForegroundPatternColorIndex, "BackgroundPatternColorIndex": local_BackgroundPatternColorIndex, "Texture": local_Texture, "ForegroundPatternColor": local_ForegroundPatternColor, "BackgroundPatternColor": local_BackgroundPatternColor, }
	if (propertyName == "Uniform"):
		retVal = this_Table.Uniform
		return retVal
	if (propertyName == "AutoFormatType"):
		retVal = this_Table.AutoFormatType
		return retVal
	if (propertyName == "Tables"):
		retVal = this_Table.Tables
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Tables", "Count": local_Count, "NestingLevel": local_NestingLevel, }
	if (propertyName == "NestingLevel"):
		retVal = this_Table.NestingLevel
		return retVal
	if (propertyName == "AllowPageBreaks"):
		retVal = this_Table.AllowPageBreaks
		return retVal
	if (propertyName == "AllowAutoFit"):
		retVal = this_Table.AllowAutoFit
		return retVal
	if (propertyName == "PreferredWidth"):
		retVal = this_Table.PreferredWidth
		return retVal
	if (propertyName == "PreferredWidthType"):
		retVal = this_Table.PreferredWidthType
		return retVal
	if (propertyName == "TopPadding"):
		retVal = this_Table.TopPadding
		return retVal
	if (propertyName == "BottomPadding"):
		retVal = this_Table.BottomPadding
		return retVal
	if (propertyName == "LeftPadding"):
		retVal = this_Table.LeftPadding
		return retVal
	if (propertyName == "RightPadding"):
		retVal = this_Table.RightPadding
		return retVal
	if (propertyName == "Spacing"):
		retVal = this_Table.Spacing
		return retVal
	if (propertyName == "TableDirection"):
		retVal = this_Table.TableDirection
		return retVal
	if (propertyName == "ID"):
		retVal = this_Table.ID
		return retVal
	if (propertyName == "Style"):
		retVal = this_Table.Style
		return retVal
	if (propertyName == "ApplyStyleHeadingRows"):
		retVal = this_Table.ApplyStyleHeadingRows
		return retVal
	if (propertyName == "ApplyStyleLastRow"):
		retVal = this_Table.ApplyStyleLastRow
		return retVal
	if (propertyName == "ApplyStyleFirstColumn"):
		retVal = this_Table.ApplyStyleFirstColumn
		return retVal
	if (propertyName == "ApplyStyleLastColumn"):
		retVal = this_Table.ApplyStyleLastColumn
		return retVal
	if (propertyName == "ApplyStyleRowBands"):
		retVal = this_Table.ApplyStyleRowBands
		return retVal
	if (propertyName == "ApplyStyleColumnBands"):
		retVal = this_Table.ApplyStyleColumnBands
		return retVal
	if (propertyName == "Title"):
		retVal = this_Table.Title
		return retVal
	if (propertyName == "Descr"):
		retVal = this_Table.Descr
		return retVal


# Tool: 338
@mcp.tool()
async def word_Table_set_Property(this_Table_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Table
	
	propertyName: Name of the property. Can be one of ...
		Borders, AllowPageBreaks, AllowAutoFit, PreferredWidth, PreferredWidthType, TopPadding, BottomPadding, LeftPadding, RightPadding, Spacing, TableDirection, ID, Style, ApplyStyleHeadingRows, ApplyStyleLastRow, ApplyStyleFirstColumn, ApplyStyleLastColumn, ApplyStyleRowBands, ApplyStyleColumnBands, Title, Descr
	"""
	this_Table = get_object(this_Table_wordObjId)
	
	EnsureWord()
	if (propertyName == "Borders"):
		this_Table.Borders = propertyValue
	if (propertyName == "AllowPageBreaks"):
		this_Table.AllowPageBreaks = propertyValue
	if (propertyName == "AllowAutoFit"):
		this_Table.AllowAutoFit = propertyValue
	if (propertyName == "PreferredWidth"):
		this_Table.PreferredWidth = propertyValue
	if (propertyName == "PreferredWidthType"):
		this_Table.PreferredWidthType = propertyValue
	if (propertyName == "TopPadding"):
		this_Table.TopPadding = propertyValue
	if (propertyName == "BottomPadding"):
		this_Table.BottomPadding = propertyValue
	if (propertyName == "LeftPadding"):
		this_Table.LeftPadding = propertyValue
	if (propertyName == "RightPadding"):
		this_Table.RightPadding = propertyValue
	if (propertyName == "Spacing"):
		this_Table.Spacing = propertyValue
	if (propertyName == "TableDirection"):
		this_Table.TableDirection = propertyValue
	if (propertyName == "ID"):
		this_Table.ID = propertyValue
	if (propertyName == "Style"):
		this_Table.Style = propertyValue
	if (propertyName == "ApplyStyleHeadingRows"):
		this_Table.ApplyStyleHeadingRows = propertyValue
	if (propertyName == "ApplyStyleLastRow"):
		this_Table.ApplyStyleLastRow = propertyValue
	if (propertyName == "ApplyStyleFirstColumn"):
		this_Table.ApplyStyleFirstColumn = propertyValue
	if (propertyName == "ApplyStyleLastColumn"):
		this_Table.ApplyStyleLastColumn = propertyValue
	if (propertyName == "ApplyStyleRowBands"):
		this_Table.ApplyStyleRowBands = propertyValue
	if (propertyName == "ApplyStyleColumnBands"):
		this_Table.ApplyStyleColumnBands = propertyValue
	if (propertyName == "Title"):
		this_Table.Title = propertyValue
	if (propertyName == "Descr"):
		this_Table.Descr = propertyValue


# Tool: 339
@mcp.tool()
async def word_Row_Select(this_Row_wordObjId: str):
	"""This tool calls the Select methodon an Row object. Pass the __WordObjectId of Row of the object you want to call the method on as the first parameter
"""
	this_Row = get_object(this_Row_wordObjId)
	this_Row.Select()


# Tool: 340
@mcp.tool()
async def word_Row_Delete(this_Row_wordObjId: str):
	"""This tool calls the Delete methodon an Row object. Pass the __WordObjectId of Row of the object you want to call the method on as the first parameter
"""
	this_Row = get_object(this_Row_wordObjId)
	this_Row.Delete()


# Tool: 341
@mcp.tool()
async def word_Row_SetLeftIndent(this_Row_wordObjId: str, LeftIndent: float, RulerStyle: int):
	"""This tool calls the SetLeftIndent methodon an Row object. Pass the __WordObjectId of Row of the object you want to call the method on as the first parameter
	
	Parameters:
		LeftIndent: the LeftIndent as float
		RulerStyle: the RulerStyle as WdRulerStyle
	"""
	this_Row = get_object(this_Row_wordObjId)
	this_Row.SetLeftIndent(LeftIndent, RulerStyle)


# Tool: 342
@mcp.tool()
async def word_Row_SetHeight(this_Row_wordObjId: str, RowHeight: float, HeightRule: int):
	"""This tool calls the SetHeight methodon an Row object. Pass the __WordObjectId of Row of the object you want to call the method on as the first parameter
	
	Parameters:
		RowHeight: the RowHeight as float
		HeightRule: the HeightRule as WdRowHeightRule
	"""
	this_Row = get_object(this_Row_wordObjId)
	this_Row.SetHeight(RowHeight, HeightRule)


# Tool: 343
@mcp.tool()
async def word_Row_ConvertToTextOld(this_Row_wordObjId: str, Separator):
	"""This tool calls the ConvertToTextOld methodon an Row object. Pass the __WordObjectId of Row of the object you want to call the method on as the first parameter
	
	Parameters:
		Separator: the Separator as VT_VARIANT
	"""
	this_Row = get_object(this_Row_wordObjId)
	Separator = tryParseString(Separator)
	retVal = this_Row.ConvertToTextOld(Separator)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 344
@mcp.tool()
async def word_Row_ConvertToText(this_Row_wordObjId: str, Separator, NestedTables):
	"""This tool calls the ConvertToText methodon an Row object. Pass the __WordObjectId of Row of the object you want to call the method on as the first parameter
	
	Parameters:
		Separator: the Separator as VT_VARIANT
		NestedTables: the NestedTables as VT_VARIANT
	"""
	this_Row = get_object(this_Row_wordObjId)
	Separator = tryParseString(Separator)
	NestedTables = tryParseString(NestedTables)
	retVal = this_Row.ConvertToText(Separator, NestedTables)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 345
@mcp.tool()
async def word_Row_get_Property(this_Row_wordObjId: str, propertyName: str):
	"""Gets properties of Row
	
	propertyName: Name of the property. Can be one of ...
		Range, AllowBreakAcrossPages, Alignment, HeadingFormat, SpaceBetweenColumns, Height, HeightRule, LeftIndent, IsLast, IsFirst, Index, Cells, Borders, Shading, Next, Previous, NestingLevel, ID
	"""
	this_Row = get_object(this_Row_wordObjId)
	
	EnsureWord()
	if (propertyName == "Range"):
		retVal = this_Row.Range
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_Start = retVal.Start
		except:
			local_Start = None
		try:
			local_End = retVal.End
		except:
			local_End = None
		try:
			local_StoryType = retVal.StoryType
		except:
			local_StoryType = None
		try:
			local_Bold = retVal.Bold
		except:
			local_Bold = None
		try:
			local_Italic = retVal.Italic
		except:
			local_Italic = None
		try:
			local_Underline = retVal.Underline
		except:
			local_Underline = None
		try:
			local_EmphasisMark = retVal.EmphasisMark
		except:
			local_EmphasisMark = None
		try:
			local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
		except:
			local_DisableCharacterSpaceGrid = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_StoryLength = retVal.StoryLength
		except:
			local_StoryLength = None
		try:
			local_LanguageID = retVal.LanguageID
		except:
			local_LanguageID = None
		try:
			local_GrammarChecked = retVal.GrammarChecked
		except:
			local_GrammarChecked = None
		try:
			local_SpellingChecked = retVal.SpellingChecked
		except:
			local_SpellingChecked = None
		try:
			local_HighlightColorIndex = retVal.HighlightColorIndex
		except:
			local_HighlightColorIndex = None
		try:
			local_CanEdit = retVal.CanEdit
		except:
			local_CanEdit = None
		try:
			local_CanPaste = retVal.CanPaste
		except:
			local_CanPaste = None
		try:
			local_IsEndOfRowMark = retVal.IsEndOfRowMark
		except:
			local_IsEndOfRowMark = None
		try:
			local_BookmarkID = retVal.BookmarkID
		except:
			local_BookmarkID = None
		try:
			local_PreviousBookmarkID = retVal.PreviousBookmarkID
		except:
			local_PreviousBookmarkID = None
		try:
			local_Case = retVal.Case
		except:
			local_Case = None
		try:
			local_Information = retVal.Information
		except:
			local_Information = None
		try:
			local_Orientation = retVal.Orientation
		except:
			local_Orientation = None
		try:
			local_LanguageIDFarEast = retVal.LanguageIDFarEast
		except:
			local_LanguageIDFarEast = None
		try:
			local_LanguageIDOther = retVal.LanguageIDOther
		except:
			local_LanguageIDOther = None
		try:
			local_LanguageDetected = retVal.LanguageDetected
		except:
			local_LanguageDetected = None
		try:
			local_FitTextWidth = retVal.FitTextWidth
		except:
			local_FitTextWidth = None
		try:
			local_HorizontalInVertical = retVal.HorizontalInVertical
		except:
			local_HorizontalInVertical = None
		try:
			local_TwoLinesInOne = retVal.TwoLinesInOne
		except:
			local_TwoLinesInOne = None
		try:
			local_CombineCharacters = retVal.CombineCharacters
		except:
			local_CombineCharacters = None
		try:
			local_NoProofing = retVal.NoProofing
		except:
			local_NoProofing = None
		try:
			local_CharacterWidth = retVal.CharacterWidth
		except:
			local_CharacterWidth = None
		try:
			local_Kana = retVal.Kana
		except:
			local_Kana = None
		try:
			local_BoldBi = retVal.BoldBi
		except:
			local_BoldBi = None
		try:
			local_ItalicBi = retVal.ItalicBi
		except:
			local_ItalicBi = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_ShowAll = retVal.ShowAll
		except:
			local_ShowAll = None
		try:
			local_CharacterStyle = retVal.CharacterStyle
		except:
			local_CharacterStyle = None
		try:
			local_ParagraphStyle = retVal.ParagraphStyle
		except:
			local_ParagraphStyle = None
		try:
			local_ListStyle = retVal.ListStyle
		except:
			local_ListStyle = None
		try:
			local_TableStyle = retVal.TableStyle
		except:
			local_TableStyle = None
		try:
			local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
		except:
			local_TextVisibleOnScreen = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }
	if (propertyName == "AllowBreakAcrossPages"):
		retVal = this_Row.AllowBreakAcrossPages
		return retVal
	if (propertyName == "Alignment"):
		retVal = this_Row.Alignment
		return retVal
	if (propertyName == "HeadingFormat"):
		retVal = this_Row.HeadingFormat
		return retVal
	if (propertyName == "SpaceBetweenColumns"):
		retVal = this_Row.SpaceBetweenColumns
		return retVal
	if (propertyName == "Height"):
		retVal = this_Row.Height
		return retVal
	if (propertyName == "HeightRule"):
		retVal = this_Row.HeightRule
		return retVal
	if (propertyName == "LeftIndent"):
		retVal = this_Row.LeftIndent
		return retVal
	if (propertyName == "IsLast"):
		retVal = this_Row.IsLast
		return retVal
	if (propertyName == "IsFirst"):
		retVal = this_Row.IsFirst
		return retVal
	if (propertyName == "Index"):
		retVal = this_Row.Index
		return retVal
	if (propertyName == "Cells"):
		retVal = this_Row.Cells
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_VerticalAlignment = retVal.VerticalAlignment
		except:
			local_VerticalAlignment = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Cells", "Count": local_Count, "Width": local_Width, "Height": local_Height, "HeightRule": local_HeightRule, "VerticalAlignment": local_VerticalAlignment, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "Borders"):
		retVal = this_Row.Borders
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Enable = retVal.Enable
		except:
			local_Enable = None
		try:
			local_DistanceFromTop = retVal.DistanceFromTop
		except:
			local_DistanceFromTop = None
		try:
			local_Shadow = retVal.Shadow
		except:
			local_Shadow = None
		try:
			local_InsideLineStyle = retVal.InsideLineStyle
		except:
			local_InsideLineStyle = None
		try:
			local_OutsideLineStyle = retVal.OutsideLineStyle
		except:
			local_OutsideLineStyle = None
		try:
			local_InsideLineWidth = retVal.InsideLineWidth
		except:
			local_InsideLineWidth = None
		try:
			local_OutsideLineWidth = retVal.OutsideLineWidth
		except:
			local_OutsideLineWidth = None
		try:
			local_InsideColorIndex = retVal.InsideColorIndex
		except:
			local_InsideColorIndex = None
		try:
			local_OutsideColorIndex = retVal.OutsideColorIndex
		except:
			local_OutsideColorIndex = None
		try:
			local_DistanceFromLeft = retVal.DistanceFromLeft
		except:
			local_DistanceFromLeft = None
		try:
			local_DistanceFromBottom = retVal.DistanceFromBottom
		except:
			local_DistanceFromBottom = None
		try:
			local_DistanceFromRight = retVal.DistanceFromRight
		except:
			local_DistanceFromRight = None
		try:
			local_AlwaysInFront = retVal.AlwaysInFront
		except:
			local_AlwaysInFront = None
		try:
			local_SurroundHeader = retVal.SurroundHeader
		except:
			local_SurroundHeader = None
		try:
			local_SurroundFooter = retVal.SurroundFooter
		except:
			local_SurroundFooter = None
		try:
			local_JoinBorders = retVal.JoinBorders
		except:
			local_JoinBorders = None
		try:
			local_HasHorizontal = retVal.HasHorizontal
		except:
			local_HasHorizontal = None
		try:
			local_HasVertical = retVal.HasVertical
		except:
			local_HasVertical = None
		try:
			local_DistanceFrom = retVal.DistanceFrom
		except:
			local_DistanceFrom = None
		try:
			local_EnableFirstPageInSection = retVal.EnableFirstPageInSection
		except:
			local_EnableFirstPageInSection = None
		try:
			local_EnableOtherPagesInSection = retVal.EnableOtherPagesInSection
		except:
			local_EnableOtherPagesInSection = None
		try:
			local_InsideColor = retVal.InsideColor
		except:
			local_InsideColor = None
		try:
			local_OutsideColor = retVal.OutsideColor
		except:
			local_OutsideColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Borders", "Count": local_Count, "Enable": local_Enable, "DistanceFromTop": local_DistanceFromTop, "Shadow": local_Shadow, "InsideLineStyle": local_InsideLineStyle, "OutsideLineStyle": local_OutsideLineStyle, "InsideLineWidth": local_InsideLineWidth, "OutsideLineWidth": local_OutsideLineWidth, "InsideColorIndex": local_InsideColorIndex, "OutsideColorIndex": local_OutsideColorIndex, "DistanceFromLeft": local_DistanceFromLeft, "DistanceFromBottom": local_DistanceFromBottom, "DistanceFromRight": local_DistanceFromRight, "AlwaysInFront": local_AlwaysInFront, "SurroundHeader": local_SurroundHeader, "SurroundFooter": local_SurroundFooter, "JoinBorders": local_JoinBorders, "HasHorizontal": local_HasHorizontal, "HasVertical": local_HasVertical, "DistanceFrom": local_DistanceFrom, "EnableFirstPageInSection": local_EnableFirstPageInSection, "EnableOtherPagesInSection": local_EnableOtherPagesInSection, "InsideColor": local_InsideColor, "OutsideColor": local_OutsideColor, }
	if (propertyName == "Shading"):
		retVal = this_Row.Shading
		try:
			local_ForegroundPatternColorIndex = retVal.ForegroundPatternColorIndex
		except:
			local_ForegroundPatternColorIndex = None
		try:
			local_BackgroundPatternColorIndex = retVal.BackgroundPatternColorIndex
		except:
			local_BackgroundPatternColorIndex = None
		try:
			local_Texture = retVal.Texture
		except:
			local_Texture = None
		try:
			local_ForegroundPatternColor = retVal.ForegroundPatternColor
		except:
			local_ForegroundPatternColor = None
		try:
			local_BackgroundPatternColor = retVal.BackgroundPatternColor
		except:
			local_BackgroundPatternColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Shading", "ForegroundPatternColorIndex": local_ForegroundPatternColorIndex, "BackgroundPatternColorIndex": local_BackgroundPatternColorIndex, "Texture": local_Texture, "ForegroundPatternColor": local_ForegroundPatternColor, "BackgroundPatternColor": local_BackgroundPatternColor, }
	if (propertyName == "Next"):
		retVal = this_Row.Next
		try:
			local_AllowBreakAcrossPages = retVal.AllowBreakAcrossPages
		except:
			local_AllowBreakAcrossPages = None
		try:
			local_Alignment = retVal.Alignment
		except:
			local_Alignment = None
		try:
			local_HeadingFormat = retVal.HeadingFormat
		except:
			local_HeadingFormat = None
		try:
			local_SpaceBetweenColumns = retVal.SpaceBetweenColumns
		except:
			local_SpaceBetweenColumns = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_LeftIndent = retVal.LeftIndent
		except:
			local_LeftIndent = None
		try:
			local_IsLast = retVal.IsLast
		except:
			local_IsLast = None
		try:
			local_IsFirst = retVal.IsFirst
		except:
			local_IsFirst = None
		try:
			local_Index = retVal.Index
		except:
			local_Index = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Row", "AllowBreakAcrossPages": local_AllowBreakAcrossPages, "Alignment": local_Alignment, "HeadingFormat": local_HeadingFormat, "SpaceBetweenColumns": local_SpaceBetweenColumns, "Height": local_Height, "HeightRule": local_HeightRule, "LeftIndent": local_LeftIndent, "IsLast": local_IsLast, "IsFirst": local_IsFirst, "Index": local_Index, "NestingLevel": local_NestingLevel, "ID": local_ID, }
	if (propertyName == "Previous"):
		retVal = this_Row.Previous
		try:
			local_AllowBreakAcrossPages = retVal.AllowBreakAcrossPages
		except:
			local_AllowBreakAcrossPages = None
		try:
			local_Alignment = retVal.Alignment
		except:
			local_Alignment = None
		try:
			local_HeadingFormat = retVal.HeadingFormat
		except:
			local_HeadingFormat = None
		try:
			local_SpaceBetweenColumns = retVal.SpaceBetweenColumns
		except:
			local_SpaceBetweenColumns = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_LeftIndent = retVal.LeftIndent
		except:
			local_LeftIndent = None
		try:
			local_IsLast = retVal.IsLast
		except:
			local_IsLast = None
		try:
			local_IsFirst = retVal.IsFirst
		except:
			local_IsFirst = None
		try:
			local_Index = retVal.Index
		except:
			local_Index = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Row", "AllowBreakAcrossPages": local_AllowBreakAcrossPages, "Alignment": local_Alignment, "HeadingFormat": local_HeadingFormat, "SpaceBetweenColumns": local_SpaceBetweenColumns, "Height": local_Height, "HeightRule": local_HeightRule, "LeftIndent": local_LeftIndent, "IsLast": local_IsLast, "IsFirst": local_IsFirst, "Index": local_Index, "NestingLevel": local_NestingLevel, "ID": local_ID, }
	if (propertyName == "NestingLevel"):
		retVal = this_Row.NestingLevel
		return retVal
	if (propertyName == "ID"):
		retVal = this_Row.ID
		return retVal


# Tool: 346
@mcp.tool()
async def word_Row_set_Property(this_Row_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Row
	
	propertyName: Name of the property. Can be one of ...
		AllowBreakAcrossPages, Alignment, HeadingFormat, SpaceBetweenColumns, Height, HeightRule, LeftIndent, Borders, ID
	"""
	this_Row = get_object(this_Row_wordObjId)
	
	EnsureWord()
	if (propertyName == "AllowBreakAcrossPages"):
		this_Row.AllowBreakAcrossPages = propertyValue
	if (propertyName == "Alignment"):
		this_Row.Alignment = propertyValue
	if (propertyName == "HeadingFormat"):
		this_Row.HeadingFormat = propertyValue
	if (propertyName == "SpaceBetweenColumns"):
		this_Row.SpaceBetweenColumns = propertyValue
	if (propertyName == "Height"):
		this_Row.Height = propertyValue
	if (propertyName == "HeightRule"):
		this_Row.HeightRule = propertyValue
	if (propertyName == "LeftIndent"):
		this_Row.LeftIndent = propertyValue
	if (propertyName == "Borders"):
		this_Row.Borders = propertyValue
	if (propertyName == "ID"):
		this_Row.ID = propertyValue


# Tool: 347
@mcp.tool()
async def word_Column_Select(this_Column_wordObjId: str):
	"""This tool calls the Select methodon an Column object. Pass the __WordObjectId of Column of the object you want to call the method on as the first parameter
"""
	this_Column = get_object(this_Column_wordObjId)
	this_Column.Select()


# Tool: 348
@mcp.tool()
async def word_Column_Delete(this_Column_wordObjId: str):
	"""This tool calls the Delete methodon an Column object. Pass the __WordObjectId of Column of the object you want to call the method on as the first parameter
"""
	this_Column = get_object(this_Column_wordObjId)
	this_Column.Delete()


# Tool: 349
@mcp.tool()
async def word_Column_SetWidth(this_Column_wordObjId: str, ColumnWidth: float, RulerStyle: int):
	"""This tool calls the SetWidth methodon an Column object. Pass the __WordObjectId of Column of the object you want to call the method on as the first parameter
	
	Parameters:
		ColumnWidth: the ColumnWidth as float
		RulerStyle: the RulerStyle as WdRulerStyle
	"""
	this_Column = get_object(this_Column_wordObjId)
	this_Column.SetWidth(ColumnWidth, RulerStyle)


# Tool: 350
@mcp.tool()
async def word_Column_AutoFit(this_Column_wordObjId: str):
	"""This tool calls the AutoFit methodon an Column object. Pass the __WordObjectId of Column of the object you want to call the method on as the first parameter
"""
	this_Column = get_object(this_Column_wordObjId)
	this_Column.AutoFit()


# Tool: 351
@mcp.tool()
async def word_Column_SortOld(this_Column_wordObjId: str, ExcludeHeader, SortFieldType, SortOrder, CaseSensitive, LanguageID):
	"""This tool calls the SortOld methodon an Column object. Pass the __WordObjectId of Column of the object you want to call the method on as the first parameter
	
	Parameters:
		ExcludeHeader: the ExcludeHeader as VT_VARIANT
		SortFieldType: the SortFieldType as VT_VARIANT
		SortOrder: the SortOrder as VT_VARIANT
		CaseSensitive: the CaseSensitive as VT_VARIANT
		LanguageID: the LanguageID as VT_VARIANT
	"""
	this_Column = get_object(this_Column_wordObjId)
	ExcludeHeader = tryParseString(ExcludeHeader)
	SortFieldType = tryParseString(SortFieldType)
	SortOrder = tryParseString(SortOrder)
	CaseSensitive = tryParseString(CaseSensitive)
	LanguageID = tryParseString(LanguageID)
	this_Column.SortOld(ExcludeHeader, SortFieldType, SortOrder, CaseSensitive, LanguageID)


# Tool: 352
@mcp.tool()
async def word_Column_Sort(this_Column_wordObjId: str, ExcludeHeader, SortFieldType, SortOrder, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID):
	"""This tool calls the Sort methodon an Column object. Pass the __WordObjectId of Column of the object you want to call the method on as the first parameter
	
	Parameters:
		ExcludeHeader: the ExcludeHeader as VT_VARIANT
		SortFieldType: the SortFieldType as VT_VARIANT
		SortOrder: the SortOrder as VT_VARIANT
		CaseSensitive: the CaseSensitive as VT_VARIANT
		BidiSort: the BidiSort as VT_VARIANT
		IgnoreThe: the IgnoreThe as VT_VARIANT
		IgnoreKashida: the IgnoreKashida as VT_VARIANT
		IgnoreDiacritics: the IgnoreDiacritics as VT_VARIANT
		IgnoreHe: the IgnoreHe as VT_VARIANT
		LanguageID: the LanguageID as VT_VARIANT
	"""
	this_Column = get_object(this_Column_wordObjId)
	ExcludeHeader = tryParseString(ExcludeHeader)
	SortFieldType = tryParseString(SortFieldType)
	SortOrder = tryParseString(SortOrder)
	CaseSensitive = tryParseString(CaseSensitive)
	BidiSort = tryParseString(BidiSort)
	IgnoreThe = tryParseString(IgnoreThe)
	IgnoreKashida = tryParseString(IgnoreKashida)
	IgnoreDiacritics = tryParseString(IgnoreDiacritics)
	IgnoreHe = tryParseString(IgnoreHe)
	LanguageID = tryParseString(LanguageID)
	this_Column.Sort(ExcludeHeader, SortFieldType, SortOrder, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID)


# Tool: 353
@mcp.tool()
async def word_Column_get_Property(this_Column_wordObjId: str, propertyName: str):
	"""Gets properties of Column
	
	propertyName: Name of the property. Can be one of ...
		Width, IsFirst, IsLast, Index, Cells, Borders, Shading, Next, Previous, NestingLevel, PreferredWidth, PreferredWidthType
	"""
	this_Column = get_object(this_Column_wordObjId)
	
	EnsureWord()
	if (propertyName == "Width"):
		retVal = this_Column.Width
		return retVal
	if (propertyName == "IsFirst"):
		retVal = this_Column.IsFirst
		return retVal
	if (propertyName == "IsLast"):
		retVal = this_Column.IsLast
		return retVal
	if (propertyName == "Index"):
		retVal = this_Column.Index
		return retVal
	if (propertyName == "Cells"):
		retVal = this_Column.Cells
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_VerticalAlignment = retVal.VerticalAlignment
		except:
			local_VerticalAlignment = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Cells", "Count": local_Count, "Width": local_Width, "Height": local_Height, "HeightRule": local_HeightRule, "VerticalAlignment": local_VerticalAlignment, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "Borders"):
		retVal = this_Column.Borders
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Enable = retVal.Enable
		except:
			local_Enable = None
		try:
			local_DistanceFromTop = retVal.DistanceFromTop
		except:
			local_DistanceFromTop = None
		try:
			local_Shadow = retVal.Shadow
		except:
			local_Shadow = None
		try:
			local_InsideLineStyle = retVal.InsideLineStyle
		except:
			local_InsideLineStyle = None
		try:
			local_OutsideLineStyle = retVal.OutsideLineStyle
		except:
			local_OutsideLineStyle = None
		try:
			local_InsideLineWidth = retVal.InsideLineWidth
		except:
			local_InsideLineWidth = None
		try:
			local_OutsideLineWidth = retVal.OutsideLineWidth
		except:
			local_OutsideLineWidth = None
		try:
			local_InsideColorIndex = retVal.InsideColorIndex
		except:
			local_InsideColorIndex = None
		try:
			local_OutsideColorIndex = retVal.OutsideColorIndex
		except:
			local_OutsideColorIndex = None
		try:
			local_DistanceFromLeft = retVal.DistanceFromLeft
		except:
			local_DistanceFromLeft = None
		try:
			local_DistanceFromBottom = retVal.DistanceFromBottom
		except:
			local_DistanceFromBottom = None
		try:
			local_DistanceFromRight = retVal.DistanceFromRight
		except:
			local_DistanceFromRight = None
		try:
			local_AlwaysInFront = retVal.AlwaysInFront
		except:
			local_AlwaysInFront = None
		try:
			local_SurroundHeader = retVal.SurroundHeader
		except:
			local_SurroundHeader = None
		try:
			local_SurroundFooter = retVal.SurroundFooter
		except:
			local_SurroundFooter = None
		try:
			local_JoinBorders = retVal.JoinBorders
		except:
			local_JoinBorders = None
		try:
			local_HasHorizontal = retVal.HasHorizontal
		except:
			local_HasHorizontal = None
		try:
			local_HasVertical = retVal.HasVertical
		except:
			local_HasVertical = None
		try:
			local_DistanceFrom = retVal.DistanceFrom
		except:
			local_DistanceFrom = None
		try:
			local_EnableFirstPageInSection = retVal.EnableFirstPageInSection
		except:
			local_EnableFirstPageInSection = None
		try:
			local_EnableOtherPagesInSection = retVal.EnableOtherPagesInSection
		except:
			local_EnableOtherPagesInSection = None
		try:
			local_InsideColor = retVal.InsideColor
		except:
			local_InsideColor = None
		try:
			local_OutsideColor = retVal.OutsideColor
		except:
			local_OutsideColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Borders", "Count": local_Count, "Enable": local_Enable, "DistanceFromTop": local_DistanceFromTop, "Shadow": local_Shadow, "InsideLineStyle": local_InsideLineStyle, "OutsideLineStyle": local_OutsideLineStyle, "InsideLineWidth": local_InsideLineWidth, "OutsideLineWidth": local_OutsideLineWidth, "InsideColorIndex": local_InsideColorIndex, "OutsideColorIndex": local_OutsideColorIndex, "DistanceFromLeft": local_DistanceFromLeft, "DistanceFromBottom": local_DistanceFromBottom, "DistanceFromRight": local_DistanceFromRight, "AlwaysInFront": local_AlwaysInFront, "SurroundHeader": local_SurroundHeader, "SurroundFooter": local_SurroundFooter, "JoinBorders": local_JoinBorders, "HasHorizontal": local_HasHorizontal, "HasVertical": local_HasVertical, "DistanceFrom": local_DistanceFrom, "EnableFirstPageInSection": local_EnableFirstPageInSection, "EnableOtherPagesInSection": local_EnableOtherPagesInSection, "InsideColor": local_InsideColor, "OutsideColor": local_OutsideColor, }
	if (propertyName == "Shading"):
		retVal = this_Column.Shading
		try:
			local_ForegroundPatternColorIndex = retVal.ForegroundPatternColorIndex
		except:
			local_ForegroundPatternColorIndex = None
		try:
			local_BackgroundPatternColorIndex = retVal.BackgroundPatternColorIndex
		except:
			local_BackgroundPatternColorIndex = None
		try:
			local_Texture = retVal.Texture
		except:
			local_Texture = None
		try:
			local_ForegroundPatternColor = retVal.ForegroundPatternColor
		except:
			local_ForegroundPatternColor = None
		try:
			local_BackgroundPatternColor = retVal.BackgroundPatternColor
		except:
			local_BackgroundPatternColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Shading", "ForegroundPatternColorIndex": local_ForegroundPatternColorIndex, "BackgroundPatternColorIndex": local_BackgroundPatternColorIndex, "Texture": local_Texture, "ForegroundPatternColor": local_ForegroundPatternColor, "BackgroundPatternColor": local_BackgroundPatternColor, }
	if (propertyName == "Next"):
		retVal = this_Column.Next
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_IsFirst = retVal.IsFirst
		except:
			local_IsFirst = None
		try:
			local_IsLast = retVal.IsLast
		except:
			local_IsLast = None
		try:
			local_Index = retVal.Index
		except:
			local_Index = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Column", "Width": local_Width, "IsFirst": local_IsFirst, "IsLast": local_IsLast, "Index": local_Index, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "Previous"):
		retVal = this_Column.Previous
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_IsFirst = retVal.IsFirst
		except:
			local_IsFirst = None
		try:
			local_IsLast = retVal.IsLast
		except:
			local_IsLast = None
		try:
			local_Index = retVal.Index
		except:
			local_Index = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Column", "Width": local_Width, "IsFirst": local_IsFirst, "IsLast": local_IsLast, "Index": local_Index, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "NestingLevel"):
		retVal = this_Column.NestingLevel
		return retVal
	if (propertyName == "PreferredWidth"):
		retVal = this_Column.PreferredWidth
		return retVal
	if (propertyName == "PreferredWidthType"):
		retVal = this_Column.PreferredWidthType
		return retVal


# Tool: 354
@mcp.tool()
async def word_Column_set_Property(this_Column_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Column
	
	propertyName: Name of the property. Can be one of ...
		Width, Borders, PreferredWidth, PreferredWidthType
	"""
	this_Column = get_object(this_Column_wordObjId)
	
	EnsureWord()
	if (propertyName == "Width"):
		this_Column.Width = propertyValue
	if (propertyName == "Borders"):
		this_Column.Borders = propertyValue
	if (propertyName == "PreferredWidth"):
		this_Column.PreferredWidth = propertyValue
	if (propertyName == "PreferredWidthType"):
		this_Column.PreferredWidthType = propertyValue


# Tool: 355
@mcp.tool()
async def word_Cell_Select(this_Cell_wordObjId: str):
	"""This tool calls the Select methodon an Cell object. Pass the __WordObjectId of Cell of the object you want to call the method on as the first parameter
"""
	this_Cell = get_object(this_Cell_wordObjId)
	this_Cell.Select()


# Tool: 356
@mcp.tool()
async def word_Cell_Delete(this_Cell_wordObjId: str, ShiftCells):
	"""This tool calls the Delete methodon an Cell object. Pass the __WordObjectId of Cell of the object you want to call the method on as the first parameter
	
	Parameters:
		ShiftCells: the ShiftCells as VT_VARIANT
	"""
	this_Cell = get_object(this_Cell_wordObjId)
	ShiftCells = tryParseString(ShiftCells)
	this_Cell.Delete(ShiftCells)


# Tool: 357
@mcp.tool()
async def word_Cell_Formula(this_Cell_wordObjId: str, Formula, NumFormat):
	"""This tool calls the Formula methodon an Cell object. Pass the __WordObjectId of Cell of the object you want to call the method on as the first parameter
	
	Parameters:
		Formula: the Formula as VT_VARIANT
		NumFormat: the NumFormat as VT_VARIANT
	"""
	this_Cell = get_object(this_Cell_wordObjId)
	Formula = tryParseString(Formula)
	NumFormat = tryParseString(NumFormat)
	this_Cell.Formula(Formula, NumFormat)


# Tool: 358
@mcp.tool()
async def word_Cell_SetWidth(this_Cell_wordObjId: str, ColumnWidth: float, RulerStyle: int):
	"""This tool calls the SetWidth methodon an Cell object. Pass the __WordObjectId of Cell of the object you want to call the method on as the first parameter
	
	Parameters:
		ColumnWidth: the ColumnWidth as float
		RulerStyle: the RulerStyle as WdRulerStyle
	"""
	this_Cell = get_object(this_Cell_wordObjId)
	this_Cell.SetWidth(ColumnWidth, RulerStyle)


# Tool: 359
@mcp.tool()
async def word_Cell_SetHeight(this_Cell_wordObjId: str, RowHeight, HeightRule: int):
	"""This tool calls the SetHeight methodon an Cell object. Pass the __WordObjectId of Cell of the object you want to call the method on as the first parameter
	
	Parameters:
		RowHeight: the RowHeight as VT_VARIANT
		HeightRule: the HeightRule as WdRowHeightRule
	"""
	this_Cell = get_object(this_Cell_wordObjId)
	RowHeight = tryParseString(RowHeight)
	this_Cell.SetHeight(RowHeight, HeightRule)


# Tool: 360
@mcp.tool()
async def word_Cell_Merge(this_Cell_wordObjId: str, MergeTo_wordObjId: str):
	"""This tool calls the Merge methodon an Cell object. Pass the __WordObjectId of Cell of the object you want to call the method on as the first parameter
	
	Parameters:
		MergeTo_wordObjId: 		To pass this object, send in the __WordObjectId of the Cell object as was obtained from a previous return value
	"""
	this_Cell = get_object(this_Cell_wordObjId)
	MergeTo = get_object(MergeTo_wordObjId)
	this_Cell.Merge(MergeTo)


# Tool: 361
@mcp.tool()
async def word_Cell_Split(this_Cell_wordObjId: str, NumRows, NumColumns):
	"""This tool calls the Split methodon an Cell object. Pass the __WordObjectId of Cell of the object you want to call the method on as the first parameter
	
	Parameters:
		NumRows: the NumRows as VT_VARIANT
		NumColumns: the NumColumns as VT_VARIANT
	"""
	this_Cell = get_object(this_Cell_wordObjId)
	NumRows = tryParseString(NumRows)
	NumColumns = tryParseString(NumColumns)
	this_Cell.Split(NumRows, NumColumns)


# Tool: 362
@mcp.tool()
async def word_Cell_AutoSum(this_Cell_wordObjId: str):
	"""This tool calls the AutoSum methodon an Cell object. Pass the __WordObjectId of Cell of the object you want to call the method on as the first parameter
"""
	this_Cell = get_object(this_Cell_wordObjId)
	this_Cell.AutoSum()


# Tool: 363
@mcp.tool()
async def word_Cell_get_Property(this_Cell_wordObjId: str, propertyName: str):
	"""Gets properties of Cell
	
	propertyName: Name of the property. Can be one of ...
		Range, RowIndex, ColumnIndex, Width, Height, HeightRule, VerticalAlignment, Column, Row, Next, Previous, Shading, Borders, Tables, NestingLevel, WordWrap, PreferredWidth, FitText, TopPadding, BottomPadding, LeftPadding, RightPadding, ID, PreferredWidthType
	"""
	this_Cell = get_object(this_Cell_wordObjId)
	
	EnsureWord()
	if (propertyName == "Range"):
		retVal = this_Cell.Range
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_Start = retVal.Start
		except:
			local_Start = None
		try:
			local_End = retVal.End
		except:
			local_End = None
		try:
			local_StoryType = retVal.StoryType
		except:
			local_StoryType = None
		try:
			local_Bold = retVal.Bold
		except:
			local_Bold = None
		try:
			local_Italic = retVal.Italic
		except:
			local_Italic = None
		try:
			local_Underline = retVal.Underline
		except:
			local_Underline = None
		try:
			local_EmphasisMark = retVal.EmphasisMark
		except:
			local_EmphasisMark = None
		try:
			local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
		except:
			local_DisableCharacterSpaceGrid = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_StoryLength = retVal.StoryLength
		except:
			local_StoryLength = None
		try:
			local_LanguageID = retVal.LanguageID
		except:
			local_LanguageID = None
		try:
			local_GrammarChecked = retVal.GrammarChecked
		except:
			local_GrammarChecked = None
		try:
			local_SpellingChecked = retVal.SpellingChecked
		except:
			local_SpellingChecked = None
		try:
			local_HighlightColorIndex = retVal.HighlightColorIndex
		except:
			local_HighlightColorIndex = None
		try:
			local_CanEdit = retVal.CanEdit
		except:
			local_CanEdit = None
		try:
			local_CanPaste = retVal.CanPaste
		except:
			local_CanPaste = None
		try:
			local_IsEndOfRowMark = retVal.IsEndOfRowMark
		except:
			local_IsEndOfRowMark = None
		try:
			local_BookmarkID = retVal.BookmarkID
		except:
			local_BookmarkID = None
		try:
			local_PreviousBookmarkID = retVal.PreviousBookmarkID
		except:
			local_PreviousBookmarkID = None
		try:
			local_Case = retVal.Case
		except:
			local_Case = None
		try:
			local_Information = retVal.Information
		except:
			local_Information = None
		try:
			local_Orientation = retVal.Orientation
		except:
			local_Orientation = None
		try:
			local_LanguageIDFarEast = retVal.LanguageIDFarEast
		except:
			local_LanguageIDFarEast = None
		try:
			local_LanguageIDOther = retVal.LanguageIDOther
		except:
			local_LanguageIDOther = None
		try:
			local_LanguageDetected = retVal.LanguageDetected
		except:
			local_LanguageDetected = None
		try:
			local_FitTextWidth = retVal.FitTextWidth
		except:
			local_FitTextWidth = None
		try:
			local_HorizontalInVertical = retVal.HorizontalInVertical
		except:
			local_HorizontalInVertical = None
		try:
			local_TwoLinesInOne = retVal.TwoLinesInOne
		except:
			local_TwoLinesInOne = None
		try:
			local_CombineCharacters = retVal.CombineCharacters
		except:
			local_CombineCharacters = None
		try:
			local_NoProofing = retVal.NoProofing
		except:
			local_NoProofing = None
		try:
			local_CharacterWidth = retVal.CharacterWidth
		except:
			local_CharacterWidth = None
		try:
			local_Kana = retVal.Kana
		except:
			local_Kana = None
		try:
			local_BoldBi = retVal.BoldBi
		except:
			local_BoldBi = None
		try:
			local_ItalicBi = retVal.ItalicBi
		except:
			local_ItalicBi = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_ShowAll = retVal.ShowAll
		except:
			local_ShowAll = None
		try:
			local_CharacterStyle = retVal.CharacterStyle
		except:
			local_CharacterStyle = None
		try:
			local_ParagraphStyle = retVal.ParagraphStyle
		except:
			local_ParagraphStyle = None
		try:
			local_ListStyle = retVal.ListStyle
		except:
			local_ListStyle = None
		try:
			local_TableStyle = retVal.TableStyle
		except:
			local_TableStyle = None
		try:
			local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
		except:
			local_TextVisibleOnScreen = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }
	if (propertyName == "RowIndex"):
		retVal = this_Cell.RowIndex
		return retVal
	if (propertyName == "ColumnIndex"):
		retVal = this_Cell.ColumnIndex
		return retVal
	if (propertyName == "Width"):
		retVal = this_Cell.Width
		return retVal
	if (propertyName == "Height"):
		retVal = this_Cell.Height
		return retVal
	if (propertyName == "HeightRule"):
		retVal = this_Cell.HeightRule
		return retVal
	if (propertyName == "VerticalAlignment"):
		retVal = this_Cell.VerticalAlignment
		return retVal
	if (propertyName == "Column"):
		retVal = this_Cell.Column
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_IsFirst = retVal.IsFirst
		except:
			local_IsFirst = None
		try:
			local_IsLast = retVal.IsLast
		except:
			local_IsLast = None
		try:
			local_Index = retVal.Index
		except:
			local_Index = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Column", "Width": local_Width, "IsFirst": local_IsFirst, "IsLast": local_IsLast, "Index": local_Index, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "Row"):
		retVal = this_Cell.Row
		try:
			local_AllowBreakAcrossPages = retVal.AllowBreakAcrossPages
		except:
			local_AllowBreakAcrossPages = None
		try:
			local_Alignment = retVal.Alignment
		except:
			local_Alignment = None
		try:
			local_HeadingFormat = retVal.HeadingFormat
		except:
			local_HeadingFormat = None
		try:
			local_SpaceBetweenColumns = retVal.SpaceBetweenColumns
		except:
			local_SpaceBetweenColumns = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_LeftIndent = retVal.LeftIndent
		except:
			local_LeftIndent = None
		try:
			local_IsLast = retVal.IsLast
		except:
			local_IsLast = None
		try:
			local_IsFirst = retVal.IsFirst
		except:
			local_IsFirst = None
		try:
			local_Index = retVal.Index
		except:
			local_Index = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Row", "AllowBreakAcrossPages": local_AllowBreakAcrossPages, "Alignment": local_Alignment, "HeadingFormat": local_HeadingFormat, "SpaceBetweenColumns": local_SpaceBetweenColumns, "Height": local_Height, "HeightRule": local_HeightRule, "LeftIndent": local_LeftIndent, "IsLast": local_IsLast, "IsFirst": local_IsFirst, "Index": local_Index, "NestingLevel": local_NestingLevel, "ID": local_ID, }
	if (propertyName == "Next"):
		retVal = this_Cell.Next
		try:
			local_RowIndex = retVal.RowIndex
		except:
			local_RowIndex = None
		try:
			local_ColumnIndex = retVal.ColumnIndex
		except:
			local_ColumnIndex = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_VerticalAlignment = retVal.VerticalAlignment
		except:
			local_VerticalAlignment = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_WordWrap = retVal.WordWrap
		except:
			local_WordWrap = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_FitText = retVal.FitText
		except:
			local_FitText = None
		try:
			local_TopPadding = retVal.TopPadding
		except:
			local_TopPadding = None
		try:
			local_BottomPadding = retVal.BottomPadding
		except:
			local_BottomPadding = None
		try:
			local_LeftPadding = retVal.LeftPadding
		except:
			local_LeftPadding = None
		try:
			local_RightPadding = retVal.RightPadding
		except:
			local_RightPadding = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Cell", "RowIndex": local_RowIndex, "ColumnIndex": local_ColumnIndex, "Width": local_Width, "Height": local_Height, "HeightRule": local_HeightRule, "VerticalAlignment": local_VerticalAlignment, "NestingLevel": local_NestingLevel, "WordWrap": local_WordWrap, "PreferredWidth": local_PreferredWidth, "FitText": local_FitText, "TopPadding": local_TopPadding, "BottomPadding": local_BottomPadding, "LeftPadding": local_LeftPadding, "RightPadding": local_RightPadding, "ID": local_ID, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "Previous"):
		retVal = this_Cell.Previous
		try:
			local_RowIndex = retVal.RowIndex
		except:
			local_RowIndex = None
		try:
			local_ColumnIndex = retVal.ColumnIndex
		except:
			local_ColumnIndex = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_VerticalAlignment = retVal.VerticalAlignment
		except:
			local_VerticalAlignment = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_WordWrap = retVal.WordWrap
		except:
			local_WordWrap = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_FitText = retVal.FitText
		except:
			local_FitText = None
		try:
			local_TopPadding = retVal.TopPadding
		except:
			local_TopPadding = None
		try:
			local_BottomPadding = retVal.BottomPadding
		except:
			local_BottomPadding = None
		try:
			local_LeftPadding = retVal.LeftPadding
		except:
			local_LeftPadding = None
		try:
			local_RightPadding = retVal.RightPadding
		except:
			local_RightPadding = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Cell", "RowIndex": local_RowIndex, "ColumnIndex": local_ColumnIndex, "Width": local_Width, "Height": local_Height, "HeightRule": local_HeightRule, "VerticalAlignment": local_VerticalAlignment, "NestingLevel": local_NestingLevel, "WordWrap": local_WordWrap, "PreferredWidth": local_PreferredWidth, "FitText": local_FitText, "TopPadding": local_TopPadding, "BottomPadding": local_BottomPadding, "LeftPadding": local_LeftPadding, "RightPadding": local_RightPadding, "ID": local_ID, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "Shading"):
		retVal = this_Cell.Shading
		try:
			local_ForegroundPatternColorIndex = retVal.ForegroundPatternColorIndex
		except:
			local_ForegroundPatternColorIndex = None
		try:
			local_BackgroundPatternColorIndex = retVal.BackgroundPatternColorIndex
		except:
			local_BackgroundPatternColorIndex = None
		try:
			local_Texture = retVal.Texture
		except:
			local_Texture = None
		try:
			local_ForegroundPatternColor = retVal.ForegroundPatternColor
		except:
			local_ForegroundPatternColor = None
		try:
			local_BackgroundPatternColor = retVal.BackgroundPatternColor
		except:
			local_BackgroundPatternColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Shading", "ForegroundPatternColorIndex": local_ForegroundPatternColorIndex, "BackgroundPatternColorIndex": local_BackgroundPatternColorIndex, "Texture": local_Texture, "ForegroundPatternColor": local_ForegroundPatternColor, "BackgroundPatternColor": local_BackgroundPatternColor, }
	if (propertyName == "Borders"):
		retVal = this_Cell.Borders
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Enable = retVal.Enable
		except:
			local_Enable = None
		try:
			local_DistanceFromTop = retVal.DistanceFromTop
		except:
			local_DistanceFromTop = None
		try:
			local_Shadow = retVal.Shadow
		except:
			local_Shadow = None
		try:
			local_InsideLineStyle = retVal.InsideLineStyle
		except:
			local_InsideLineStyle = None
		try:
			local_OutsideLineStyle = retVal.OutsideLineStyle
		except:
			local_OutsideLineStyle = None
		try:
			local_InsideLineWidth = retVal.InsideLineWidth
		except:
			local_InsideLineWidth = None
		try:
			local_OutsideLineWidth = retVal.OutsideLineWidth
		except:
			local_OutsideLineWidth = None
		try:
			local_InsideColorIndex = retVal.InsideColorIndex
		except:
			local_InsideColorIndex = None
		try:
			local_OutsideColorIndex = retVal.OutsideColorIndex
		except:
			local_OutsideColorIndex = None
		try:
			local_DistanceFromLeft = retVal.DistanceFromLeft
		except:
			local_DistanceFromLeft = None
		try:
			local_DistanceFromBottom = retVal.DistanceFromBottom
		except:
			local_DistanceFromBottom = None
		try:
			local_DistanceFromRight = retVal.DistanceFromRight
		except:
			local_DistanceFromRight = None
		try:
			local_AlwaysInFront = retVal.AlwaysInFront
		except:
			local_AlwaysInFront = None
		try:
			local_SurroundHeader = retVal.SurroundHeader
		except:
			local_SurroundHeader = None
		try:
			local_SurroundFooter = retVal.SurroundFooter
		except:
			local_SurroundFooter = None
		try:
			local_JoinBorders = retVal.JoinBorders
		except:
			local_JoinBorders = None
		try:
			local_HasHorizontal = retVal.HasHorizontal
		except:
			local_HasHorizontal = None
		try:
			local_HasVertical = retVal.HasVertical
		except:
			local_HasVertical = None
		try:
			local_DistanceFrom = retVal.DistanceFrom
		except:
			local_DistanceFrom = None
		try:
			local_EnableFirstPageInSection = retVal.EnableFirstPageInSection
		except:
			local_EnableFirstPageInSection = None
		try:
			local_EnableOtherPagesInSection = retVal.EnableOtherPagesInSection
		except:
			local_EnableOtherPagesInSection = None
		try:
			local_InsideColor = retVal.InsideColor
		except:
			local_InsideColor = None
		try:
			local_OutsideColor = retVal.OutsideColor
		except:
			local_OutsideColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Borders", "Count": local_Count, "Enable": local_Enable, "DistanceFromTop": local_DistanceFromTop, "Shadow": local_Shadow, "InsideLineStyle": local_InsideLineStyle, "OutsideLineStyle": local_OutsideLineStyle, "InsideLineWidth": local_InsideLineWidth, "OutsideLineWidth": local_OutsideLineWidth, "InsideColorIndex": local_InsideColorIndex, "OutsideColorIndex": local_OutsideColorIndex, "DistanceFromLeft": local_DistanceFromLeft, "DistanceFromBottom": local_DistanceFromBottom, "DistanceFromRight": local_DistanceFromRight, "AlwaysInFront": local_AlwaysInFront, "SurroundHeader": local_SurroundHeader, "SurroundFooter": local_SurroundFooter, "JoinBorders": local_JoinBorders, "HasHorizontal": local_HasHorizontal, "HasVertical": local_HasVertical, "DistanceFrom": local_DistanceFrom, "EnableFirstPageInSection": local_EnableFirstPageInSection, "EnableOtherPagesInSection": local_EnableOtherPagesInSection, "InsideColor": local_InsideColor, "OutsideColor": local_OutsideColor, }
	if (propertyName == "Tables"):
		retVal = this_Cell.Tables
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Tables", "Count": local_Count, "NestingLevel": local_NestingLevel, }
	if (propertyName == "NestingLevel"):
		retVal = this_Cell.NestingLevel
		return retVal
	if (propertyName == "WordWrap"):
		retVal = this_Cell.WordWrap
		return retVal
	if (propertyName == "PreferredWidth"):
		retVal = this_Cell.PreferredWidth
		return retVal
	if (propertyName == "FitText"):
		retVal = this_Cell.FitText
		return retVal
	if (propertyName == "TopPadding"):
		retVal = this_Cell.TopPadding
		return retVal
	if (propertyName == "BottomPadding"):
		retVal = this_Cell.BottomPadding
		return retVal
	if (propertyName == "LeftPadding"):
		retVal = this_Cell.LeftPadding
		return retVal
	if (propertyName == "RightPadding"):
		retVal = this_Cell.RightPadding
		return retVal
	if (propertyName == "ID"):
		retVal = this_Cell.ID
		return retVal
	if (propertyName == "PreferredWidthType"):
		retVal = this_Cell.PreferredWidthType
		return retVal


# Tool: 364
@mcp.tool()
async def word_Cell_set_Property(this_Cell_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Cell
	
	propertyName: Name of the property. Can be one of ...
		Width, Height, HeightRule, VerticalAlignment, Borders, WordWrap, PreferredWidth, FitText, TopPadding, BottomPadding, LeftPadding, RightPadding, ID, PreferredWidthType
	"""
	this_Cell = get_object(this_Cell_wordObjId)
	
	EnsureWord()
	if (propertyName == "Width"):
		this_Cell.Width = propertyValue
	if (propertyName == "Height"):
		this_Cell.Height = propertyValue
	if (propertyName == "HeightRule"):
		this_Cell.HeightRule = propertyValue
	if (propertyName == "VerticalAlignment"):
		this_Cell.VerticalAlignment = propertyValue
	if (propertyName == "Borders"):
		this_Cell.Borders = propertyValue
	if (propertyName == "WordWrap"):
		this_Cell.WordWrap = propertyValue
	if (propertyName == "PreferredWidth"):
		this_Cell.PreferredWidth = propertyValue
	if (propertyName == "FitText"):
		this_Cell.FitText = propertyValue
	if (propertyName == "TopPadding"):
		this_Cell.TopPadding = propertyValue
	if (propertyName == "BottomPadding"):
		this_Cell.BottomPadding = propertyValue
	if (propertyName == "LeftPadding"):
		this_Cell.LeftPadding = propertyValue
	if (propertyName == "RightPadding"):
		this_Cell.RightPadding = propertyValue
	if (propertyName == "ID"):
		this_Cell.ID = propertyValue
	if (propertyName == "PreferredWidthType"):
		this_Cell.PreferredWidthType = propertyValue


# Tool: 365
@mcp.tool()
async def word_Tables_Item(this_Tables_wordObjId: str, Index: int):
	"""This tool calls the Item methodon an Tables object. Pass the __WordObjectId of Tables of the object you want to call the method on as the first parameter
	
	Parameters:
		Index: the Index as int
	"""
	this_Tables = get_object(this_Tables_wordObjId)
	retVal = this_Tables.Item(Index)
	try:
		local_Uniform = retVal.Uniform
	except:
		local_Uniform = None
	try:
		local_AutoFormatType = retVal.AutoFormatType
	except:
		local_AutoFormatType = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_AllowPageBreaks = retVal.AllowPageBreaks
	except:
		local_AllowPageBreaks = None
	try:
		local_AllowAutoFit = retVal.AllowAutoFit
	except:
		local_AllowAutoFit = None
	try:
		local_PreferredWidth = retVal.PreferredWidth
	except:
		local_PreferredWidth = None
	try:
		local_PreferredWidthType = retVal.PreferredWidthType
	except:
		local_PreferredWidthType = None
	try:
		local_TopPadding = retVal.TopPadding
	except:
		local_TopPadding = None
	try:
		local_BottomPadding = retVal.BottomPadding
	except:
		local_BottomPadding = None
	try:
		local_LeftPadding = retVal.LeftPadding
	except:
		local_LeftPadding = None
	try:
		local_RightPadding = retVal.RightPadding
	except:
		local_RightPadding = None
	try:
		local_Spacing = retVal.Spacing
	except:
		local_Spacing = None
	try:
		local_TableDirection = retVal.TableDirection
	except:
		local_TableDirection = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_ApplyStyleHeadingRows = retVal.ApplyStyleHeadingRows
	except:
		local_ApplyStyleHeadingRows = None
	try:
		local_ApplyStyleLastRow = retVal.ApplyStyleLastRow
	except:
		local_ApplyStyleLastRow = None
	try:
		local_ApplyStyleFirstColumn = retVal.ApplyStyleFirstColumn
	except:
		local_ApplyStyleFirstColumn = None
	try:
		local_ApplyStyleLastColumn = retVal.ApplyStyleLastColumn
	except:
		local_ApplyStyleLastColumn = None
	try:
		local_ApplyStyleRowBands = retVal.ApplyStyleRowBands
	except:
		local_ApplyStyleRowBands = None
	try:
		local_ApplyStyleColumnBands = retVal.ApplyStyleColumnBands
	except:
		local_ApplyStyleColumnBands = None
	try:
		local_Title = retVal.Title
	except:
		local_Title = None
	try:
		local_Descr = retVal.Descr
	except:
		local_Descr = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Table", "Uniform": local_Uniform, "AutoFormatType": local_AutoFormatType, "NestingLevel": local_NestingLevel, "AllowPageBreaks": local_AllowPageBreaks, "AllowAutoFit": local_AllowAutoFit, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, "TopPadding": local_TopPadding, "BottomPadding": local_BottomPadding, "LeftPadding": local_LeftPadding, "RightPadding": local_RightPadding, "Spacing": local_Spacing, "TableDirection": local_TableDirection, "ID": local_ID, "Style": local_Style, "ApplyStyleHeadingRows": local_ApplyStyleHeadingRows, "ApplyStyleLastRow": local_ApplyStyleLastRow, "ApplyStyleFirstColumn": local_ApplyStyleFirstColumn, "ApplyStyleLastColumn": local_ApplyStyleLastColumn, "ApplyStyleRowBands": local_ApplyStyleRowBands, "ApplyStyleColumnBands": local_ApplyStyleColumnBands, "Title": local_Title, "Descr": local_Descr, }


# Tool: 366
@mcp.tool()
async def word_Tables_AddOld(this_Tables_wordObjId: str, Range_wordObjId: str, NumRows: int, NumColumns: int):
	"""This tool calls the AddOld methodon an Tables object. Pass the __WordObjectId of Tables of the object you want to call the method on as the first parameter
	
	Parameters:
		Range_wordObjId: 		To pass this object, send in the __WordObjectId of the Range object as was obtained from a previous return value
		NumRows: the NumRows as int
		NumColumns: the NumColumns as int
	"""
	this_Tables = get_object(this_Tables_wordObjId)
	Range = get_object(Range_wordObjId)
	retVal = this_Tables.AddOld(Range, NumRows, NumColumns)
	try:
		local_Uniform = retVal.Uniform
	except:
		local_Uniform = None
	try:
		local_AutoFormatType = retVal.AutoFormatType
	except:
		local_AutoFormatType = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_AllowPageBreaks = retVal.AllowPageBreaks
	except:
		local_AllowPageBreaks = None
	try:
		local_AllowAutoFit = retVal.AllowAutoFit
	except:
		local_AllowAutoFit = None
	try:
		local_PreferredWidth = retVal.PreferredWidth
	except:
		local_PreferredWidth = None
	try:
		local_PreferredWidthType = retVal.PreferredWidthType
	except:
		local_PreferredWidthType = None
	try:
		local_TopPadding = retVal.TopPadding
	except:
		local_TopPadding = None
	try:
		local_BottomPadding = retVal.BottomPadding
	except:
		local_BottomPadding = None
	try:
		local_LeftPadding = retVal.LeftPadding
	except:
		local_LeftPadding = None
	try:
		local_RightPadding = retVal.RightPadding
	except:
		local_RightPadding = None
	try:
		local_Spacing = retVal.Spacing
	except:
		local_Spacing = None
	try:
		local_TableDirection = retVal.TableDirection
	except:
		local_TableDirection = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_ApplyStyleHeadingRows = retVal.ApplyStyleHeadingRows
	except:
		local_ApplyStyleHeadingRows = None
	try:
		local_ApplyStyleLastRow = retVal.ApplyStyleLastRow
	except:
		local_ApplyStyleLastRow = None
	try:
		local_ApplyStyleFirstColumn = retVal.ApplyStyleFirstColumn
	except:
		local_ApplyStyleFirstColumn = None
	try:
		local_ApplyStyleLastColumn = retVal.ApplyStyleLastColumn
	except:
		local_ApplyStyleLastColumn = None
	try:
		local_ApplyStyleRowBands = retVal.ApplyStyleRowBands
	except:
		local_ApplyStyleRowBands = None
	try:
		local_ApplyStyleColumnBands = retVal.ApplyStyleColumnBands
	except:
		local_ApplyStyleColumnBands = None
	try:
		local_Title = retVal.Title
	except:
		local_Title = None
	try:
		local_Descr = retVal.Descr
	except:
		local_Descr = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Table", "Uniform": local_Uniform, "AutoFormatType": local_AutoFormatType, "NestingLevel": local_NestingLevel, "AllowPageBreaks": local_AllowPageBreaks, "AllowAutoFit": local_AllowAutoFit, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, "TopPadding": local_TopPadding, "BottomPadding": local_BottomPadding, "LeftPadding": local_LeftPadding, "RightPadding": local_RightPadding, "Spacing": local_Spacing, "TableDirection": local_TableDirection, "ID": local_ID, "Style": local_Style, "ApplyStyleHeadingRows": local_ApplyStyleHeadingRows, "ApplyStyleLastRow": local_ApplyStyleLastRow, "ApplyStyleFirstColumn": local_ApplyStyleFirstColumn, "ApplyStyleLastColumn": local_ApplyStyleLastColumn, "ApplyStyleRowBands": local_ApplyStyleRowBands, "ApplyStyleColumnBands": local_ApplyStyleColumnBands, "Title": local_Title, "Descr": local_Descr, }


# Tool: 367
@mcp.tool()
async def word_Tables_Add(this_Tables_wordObjId: str, Range_wordObjId: str, NumRows: int, NumColumns: int, DefaultTableBehavior, AutoFitBehavior):
	"""This tool calls the Add methodon an Tables object. Pass the __WordObjectId of Tables of the object you want to call the method on as the first parameter
	
	Parameters:
		Range_wordObjId: 		To pass this object, send in the __WordObjectId of the Range object as was obtained from a previous return value
		NumRows: the NumRows as int
		NumColumns: the NumColumns as int
		DefaultTableBehavior: the DefaultTableBehavior as VT_VARIANT
		AutoFitBehavior: the AutoFitBehavior as VT_VARIANT
	"""
	this_Tables = get_object(this_Tables_wordObjId)
	Range = get_object(Range_wordObjId)
	DefaultTableBehavior = tryParseString(DefaultTableBehavior)
	AutoFitBehavior = tryParseString(AutoFitBehavior)
	retVal = this_Tables.Add(Range, NumRows, NumColumns, DefaultTableBehavior, AutoFitBehavior)
	try:
		local_Uniform = retVal.Uniform
	except:
		local_Uniform = None
	try:
		local_AutoFormatType = retVal.AutoFormatType
	except:
		local_AutoFormatType = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_AllowPageBreaks = retVal.AllowPageBreaks
	except:
		local_AllowPageBreaks = None
	try:
		local_AllowAutoFit = retVal.AllowAutoFit
	except:
		local_AllowAutoFit = None
	try:
		local_PreferredWidth = retVal.PreferredWidth
	except:
		local_PreferredWidth = None
	try:
		local_PreferredWidthType = retVal.PreferredWidthType
	except:
		local_PreferredWidthType = None
	try:
		local_TopPadding = retVal.TopPadding
	except:
		local_TopPadding = None
	try:
		local_BottomPadding = retVal.BottomPadding
	except:
		local_BottomPadding = None
	try:
		local_LeftPadding = retVal.LeftPadding
	except:
		local_LeftPadding = None
	try:
		local_RightPadding = retVal.RightPadding
	except:
		local_RightPadding = None
	try:
		local_Spacing = retVal.Spacing
	except:
		local_Spacing = None
	try:
		local_TableDirection = retVal.TableDirection
	except:
		local_TableDirection = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_ApplyStyleHeadingRows = retVal.ApplyStyleHeadingRows
	except:
		local_ApplyStyleHeadingRows = None
	try:
		local_ApplyStyleLastRow = retVal.ApplyStyleLastRow
	except:
		local_ApplyStyleLastRow = None
	try:
		local_ApplyStyleFirstColumn = retVal.ApplyStyleFirstColumn
	except:
		local_ApplyStyleFirstColumn = None
	try:
		local_ApplyStyleLastColumn = retVal.ApplyStyleLastColumn
	except:
		local_ApplyStyleLastColumn = None
	try:
		local_ApplyStyleRowBands = retVal.ApplyStyleRowBands
	except:
		local_ApplyStyleRowBands = None
	try:
		local_ApplyStyleColumnBands = retVal.ApplyStyleColumnBands
	except:
		local_ApplyStyleColumnBands = None
	try:
		local_Title = retVal.Title
	except:
		local_Title = None
	try:
		local_Descr = retVal.Descr
	except:
		local_Descr = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Table", "Uniform": local_Uniform, "AutoFormatType": local_AutoFormatType, "NestingLevel": local_NestingLevel, "AllowPageBreaks": local_AllowPageBreaks, "AllowAutoFit": local_AllowAutoFit, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, "TopPadding": local_TopPadding, "BottomPadding": local_BottomPadding, "LeftPadding": local_LeftPadding, "RightPadding": local_RightPadding, "Spacing": local_Spacing, "TableDirection": local_TableDirection, "ID": local_ID, "Style": local_Style, "ApplyStyleHeadingRows": local_ApplyStyleHeadingRows, "ApplyStyleLastRow": local_ApplyStyleLastRow, "ApplyStyleFirstColumn": local_ApplyStyleFirstColumn, "ApplyStyleLastColumn": local_ApplyStyleLastColumn, "ApplyStyleRowBands": local_ApplyStyleRowBands, "ApplyStyleColumnBands": local_ApplyStyleColumnBands, "Title": local_Title, "Descr": local_Descr, }


# Tool: 368
@mcp.tool()
async def word_Tables_get_Property(this_Tables_wordObjId: str, propertyName: str):
	"""Gets properties of Tables
	
	propertyName: Name of the property. Can be one of ...
		Count, NestingLevel
	"""
	this_Tables = get_object(this_Tables_wordObjId)
	
	EnsureWord()
	if (propertyName == "Count"):
		retVal = this_Tables.Count
		return retVal
	if (propertyName == "NestingLevel"):
		retVal = this_Tables.NestingLevel
		return retVal


# Tool: 369
@mcp.tool()
async def word_Tables_set_Property(this_Tables_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Tables
	
	propertyName: Name of the property. Can be one of ...
		
	"""
	this_Tables = get_object(this_Tables_wordObjId)
	
	EnsureWord()


# Tool: 370
@mcp.tool()
async def word_Rows_Item(this_Rows_wordObjId: str, Index: int):
	"""This tool calls the Item methodon an Rows object. Pass the __WordObjectId of Rows of the object you want to call the method on as the first parameter
	
	Parameters:
		Index: the Index as int
	"""
	this_Rows = get_object(this_Rows_wordObjId)
	retVal = this_Rows.Item(Index)
	try:
		local_AllowBreakAcrossPages = retVal.AllowBreakAcrossPages
	except:
		local_AllowBreakAcrossPages = None
	try:
		local_Alignment = retVal.Alignment
	except:
		local_Alignment = None
	try:
		local_HeadingFormat = retVal.HeadingFormat
	except:
		local_HeadingFormat = None
	try:
		local_SpaceBetweenColumns = retVal.SpaceBetweenColumns
	except:
		local_SpaceBetweenColumns = None
	try:
		local_Height = retVal.Height
	except:
		local_Height = None
	try:
		local_HeightRule = retVal.HeightRule
	except:
		local_HeightRule = None
	try:
		local_LeftIndent = retVal.LeftIndent
	except:
		local_LeftIndent = None
	try:
		local_IsLast = retVal.IsLast
	except:
		local_IsLast = None
	try:
		local_IsFirst = retVal.IsFirst
	except:
		local_IsFirst = None
	try:
		local_Index = retVal.Index
	except:
		local_Index = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Row", "AllowBreakAcrossPages": local_AllowBreakAcrossPages, "Alignment": local_Alignment, "HeadingFormat": local_HeadingFormat, "SpaceBetweenColumns": local_SpaceBetweenColumns, "Height": local_Height, "HeightRule": local_HeightRule, "LeftIndent": local_LeftIndent, "IsLast": local_IsLast, "IsFirst": local_IsFirst, "Index": local_Index, "NestingLevel": local_NestingLevel, "ID": local_ID, }


# Tool: 371
@mcp.tool()
async def word_Rows_Add(this_Rows_wordObjId: str, BeforeRow):
	"""This tool calls the Add methodon an Rows object. Pass the __WordObjectId of Rows of the object you want to call the method on as the first parameter
	
	Parameters:
		BeforeRow: the BeforeRow as VT_VARIANT
	"""
	this_Rows = get_object(this_Rows_wordObjId)
	BeforeRow = tryParseString(BeforeRow)
	retVal = this_Rows.Add(BeforeRow)
	try:
		local_AllowBreakAcrossPages = retVal.AllowBreakAcrossPages
	except:
		local_AllowBreakAcrossPages = None
	try:
		local_Alignment = retVal.Alignment
	except:
		local_Alignment = None
	try:
		local_HeadingFormat = retVal.HeadingFormat
	except:
		local_HeadingFormat = None
	try:
		local_SpaceBetweenColumns = retVal.SpaceBetweenColumns
	except:
		local_SpaceBetweenColumns = None
	try:
		local_Height = retVal.Height
	except:
		local_Height = None
	try:
		local_HeightRule = retVal.HeightRule
	except:
		local_HeightRule = None
	try:
		local_LeftIndent = retVal.LeftIndent
	except:
		local_LeftIndent = None
	try:
		local_IsLast = retVal.IsLast
	except:
		local_IsLast = None
	try:
		local_IsFirst = retVal.IsFirst
	except:
		local_IsFirst = None
	try:
		local_Index = retVal.Index
	except:
		local_Index = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Row", "AllowBreakAcrossPages": local_AllowBreakAcrossPages, "Alignment": local_Alignment, "HeadingFormat": local_HeadingFormat, "SpaceBetweenColumns": local_SpaceBetweenColumns, "Height": local_Height, "HeightRule": local_HeightRule, "LeftIndent": local_LeftIndent, "IsLast": local_IsLast, "IsFirst": local_IsFirst, "Index": local_Index, "NestingLevel": local_NestingLevel, "ID": local_ID, }


# Tool: 372
@mcp.tool()
async def word_Rows_Select(this_Rows_wordObjId: str):
	"""This tool calls the Select methodon an Rows object. Pass the __WordObjectId of Rows of the object you want to call the method on as the first parameter
"""
	this_Rows = get_object(this_Rows_wordObjId)
	this_Rows.Select()


# Tool: 373
@mcp.tool()
async def word_Rows_Delete(this_Rows_wordObjId: str):
	"""This tool calls the Delete methodon an Rows object. Pass the __WordObjectId of Rows of the object you want to call the method on as the first parameter
"""
	this_Rows = get_object(this_Rows_wordObjId)
	this_Rows.Delete()


# Tool: 374
@mcp.tool()
async def word_Rows_SetLeftIndent(this_Rows_wordObjId: str, LeftIndent: float, RulerStyle: int):
	"""This tool calls the SetLeftIndent methodon an Rows object. Pass the __WordObjectId of Rows of the object you want to call the method on as the first parameter
	
	Parameters:
		LeftIndent: the LeftIndent as float
		RulerStyle: the RulerStyle as WdRulerStyle
	"""
	this_Rows = get_object(this_Rows_wordObjId)
	this_Rows.SetLeftIndent(LeftIndent, RulerStyle)


# Tool: 375
@mcp.tool()
async def word_Rows_SetHeight(this_Rows_wordObjId: str, RowHeight: float, HeightRule: int):
	"""This tool calls the SetHeight methodon an Rows object. Pass the __WordObjectId of Rows of the object you want to call the method on as the first parameter
	
	Parameters:
		RowHeight: the RowHeight as float
		HeightRule: the HeightRule as WdRowHeightRule
	"""
	this_Rows = get_object(this_Rows_wordObjId)
	this_Rows.SetHeight(RowHeight, HeightRule)


# Tool: 376
@mcp.tool()
async def word_Rows_ConvertToTextOld(this_Rows_wordObjId: str, Separator):
	"""This tool calls the ConvertToTextOld methodon an Rows object. Pass the __WordObjectId of Rows of the object you want to call the method on as the first parameter
	
	Parameters:
		Separator: the Separator as VT_VARIANT
	"""
	this_Rows = get_object(this_Rows_wordObjId)
	Separator = tryParseString(Separator)
	retVal = this_Rows.ConvertToTextOld(Separator)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 377
@mcp.tool()
async def word_Rows_DistributeHeight(this_Rows_wordObjId: str):
	"""This tool calls the DistributeHeight methodon an Rows object. Pass the __WordObjectId of Rows of the object you want to call the method on as the first parameter
"""
	this_Rows = get_object(this_Rows_wordObjId)
	this_Rows.DistributeHeight()


# Tool: 378
@mcp.tool()
async def word_Rows_ConvertToText(this_Rows_wordObjId: str, Separator, NestedTables):
	"""This tool calls the ConvertToText methodon an Rows object. Pass the __WordObjectId of Rows of the object you want to call the method on as the first parameter
	
	Parameters:
		Separator: the Separator as VT_VARIANT
		NestedTables: the NestedTables as VT_VARIANT
	"""
	this_Rows = get_object(this_Rows_wordObjId)
	Separator = tryParseString(Separator)
	NestedTables = tryParseString(NestedTables)
	retVal = this_Rows.ConvertToText(Separator, NestedTables)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 379
@mcp.tool()
async def word_Rows_get_Property(this_Rows_wordObjId: str, propertyName: str):
	"""Gets properties of Rows
	
	propertyName: Name of the property. Can be one of ...
		Count, AllowBreakAcrossPages, Alignment, HeadingFormat, SpaceBetweenColumns, Height, HeightRule, LeftIndent, First, Last, Borders, Shading, WrapAroundText, DistanceTop, DistanceBottom, DistanceLeft, DistanceRight, HorizontalPosition, VerticalPosition, RelativeHorizontalPosition, RelativeVerticalPosition, AllowOverlap, NestingLevel, TableDirection
	"""
	this_Rows = get_object(this_Rows_wordObjId)
	
	EnsureWord()
	if (propertyName == "Count"):
		retVal = this_Rows.Count
		return retVal
	if (propertyName == "AllowBreakAcrossPages"):
		retVal = this_Rows.AllowBreakAcrossPages
		return retVal
	if (propertyName == "Alignment"):
		retVal = this_Rows.Alignment
		return retVal
	if (propertyName == "HeadingFormat"):
		retVal = this_Rows.HeadingFormat
		return retVal
	if (propertyName == "SpaceBetweenColumns"):
		retVal = this_Rows.SpaceBetweenColumns
		return retVal
	if (propertyName == "Height"):
		retVal = this_Rows.Height
		return retVal
	if (propertyName == "HeightRule"):
		retVal = this_Rows.HeightRule
		return retVal
	if (propertyName == "LeftIndent"):
		retVal = this_Rows.LeftIndent
		return retVal
	if (propertyName == "First"):
		retVal = this_Rows.First
		try:
			local_AllowBreakAcrossPages = retVal.AllowBreakAcrossPages
		except:
			local_AllowBreakAcrossPages = None
		try:
			local_Alignment = retVal.Alignment
		except:
			local_Alignment = None
		try:
			local_HeadingFormat = retVal.HeadingFormat
		except:
			local_HeadingFormat = None
		try:
			local_SpaceBetweenColumns = retVal.SpaceBetweenColumns
		except:
			local_SpaceBetweenColumns = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_LeftIndent = retVal.LeftIndent
		except:
			local_LeftIndent = None
		try:
			local_IsLast = retVal.IsLast
		except:
			local_IsLast = None
		try:
			local_IsFirst = retVal.IsFirst
		except:
			local_IsFirst = None
		try:
			local_Index = retVal.Index
		except:
			local_Index = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Row", "AllowBreakAcrossPages": local_AllowBreakAcrossPages, "Alignment": local_Alignment, "HeadingFormat": local_HeadingFormat, "SpaceBetweenColumns": local_SpaceBetweenColumns, "Height": local_Height, "HeightRule": local_HeightRule, "LeftIndent": local_LeftIndent, "IsLast": local_IsLast, "IsFirst": local_IsFirst, "Index": local_Index, "NestingLevel": local_NestingLevel, "ID": local_ID, }
	if (propertyName == "Last"):
		retVal = this_Rows.Last
		try:
			local_AllowBreakAcrossPages = retVal.AllowBreakAcrossPages
		except:
			local_AllowBreakAcrossPages = None
		try:
			local_Alignment = retVal.Alignment
		except:
			local_Alignment = None
		try:
			local_HeadingFormat = retVal.HeadingFormat
		except:
			local_HeadingFormat = None
		try:
			local_SpaceBetweenColumns = retVal.SpaceBetweenColumns
		except:
			local_SpaceBetweenColumns = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_LeftIndent = retVal.LeftIndent
		except:
			local_LeftIndent = None
		try:
			local_IsLast = retVal.IsLast
		except:
			local_IsLast = None
		try:
			local_IsFirst = retVal.IsFirst
		except:
			local_IsFirst = None
		try:
			local_Index = retVal.Index
		except:
			local_Index = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Row", "AllowBreakAcrossPages": local_AllowBreakAcrossPages, "Alignment": local_Alignment, "HeadingFormat": local_HeadingFormat, "SpaceBetweenColumns": local_SpaceBetweenColumns, "Height": local_Height, "HeightRule": local_HeightRule, "LeftIndent": local_LeftIndent, "IsLast": local_IsLast, "IsFirst": local_IsFirst, "Index": local_Index, "NestingLevel": local_NestingLevel, "ID": local_ID, }
	if (propertyName == "Borders"):
		retVal = this_Rows.Borders
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Enable = retVal.Enable
		except:
			local_Enable = None
		try:
			local_DistanceFromTop = retVal.DistanceFromTop
		except:
			local_DistanceFromTop = None
		try:
			local_Shadow = retVal.Shadow
		except:
			local_Shadow = None
		try:
			local_InsideLineStyle = retVal.InsideLineStyle
		except:
			local_InsideLineStyle = None
		try:
			local_OutsideLineStyle = retVal.OutsideLineStyle
		except:
			local_OutsideLineStyle = None
		try:
			local_InsideLineWidth = retVal.InsideLineWidth
		except:
			local_InsideLineWidth = None
		try:
			local_OutsideLineWidth = retVal.OutsideLineWidth
		except:
			local_OutsideLineWidth = None
		try:
			local_InsideColorIndex = retVal.InsideColorIndex
		except:
			local_InsideColorIndex = None
		try:
			local_OutsideColorIndex = retVal.OutsideColorIndex
		except:
			local_OutsideColorIndex = None
		try:
			local_DistanceFromLeft = retVal.DistanceFromLeft
		except:
			local_DistanceFromLeft = None
		try:
			local_DistanceFromBottom = retVal.DistanceFromBottom
		except:
			local_DistanceFromBottom = None
		try:
			local_DistanceFromRight = retVal.DistanceFromRight
		except:
			local_DistanceFromRight = None
		try:
			local_AlwaysInFront = retVal.AlwaysInFront
		except:
			local_AlwaysInFront = None
		try:
			local_SurroundHeader = retVal.SurroundHeader
		except:
			local_SurroundHeader = None
		try:
			local_SurroundFooter = retVal.SurroundFooter
		except:
			local_SurroundFooter = None
		try:
			local_JoinBorders = retVal.JoinBorders
		except:
			local_JoinBorders = None
		try:
			local_HasHorizontal = retVal.HasHorizontal
		except:
			local_HasHorizontal = None
		try:
			local_HasVertical = retVal.HasVertical
		except:
			local_HasVertical = None
		try:
			local_DistanceFrom = retVal.DistanceFrom
		except:
			local_DistanceFrom = None
		try:
			local_EnableFirstPageInSection = retVal.EnableFirstPageInSection
		except:
			local_EnableFirstPageInSection = None
		try:
			local_EnableOtherPagesInSection = retVal.EnableOtherPagesInSection
		except:
			local_EnableOtherPagesInSection = None
		try:
			local_InsideColor = retVal.InsideColor
		except:
			local_InsideColor = None
		try:
			local_OutsideColor = retVal.OutsideColor
		except:
			local_OutsideColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Borders", "Count": local_Count, "Enable": local_Enable, "DistanceFromTop": local_DistanceFromTop, "Shadow": local_Shadow, "InsideLineStyle": local_InsideLineStyle, "OutsideLineStyle": local_OutsideLineStyle, "InsideLineWidth": local_InsideLineWidth, "OutsideLineWidth": local_OutsideLineWidth, "InsideColorIndex": local_InsideColorIndex, "OutsideColorIndex": local_OutsideColorIndex, "DistanceFromLeft": local_DistanceFromLeft, "DistanceFromBottom": local_DistanceFromBottom, "DistanceFromRight": local_DistanceFromRight, "AlwaysInFront": local_AlwaysInFront, "SurroundHeader": local_SurroundHeader, "SurroundFooter": local_SurroundFooter, "JoinBorders": local_JoinBorders, "HasHorizontal": local_HasHorizontal, "HasVertical": local_HasVertical, "DistanceFrom": local_DistanceFrom, "EnableFirstPageInSection": local_EnableFirstPageInSection, "EnableOtherPagesInSection": local_EnableOtherPagesInSection, "InsideColor": local_InsideColor, "OutsideColor": local_OutsideColor, }
	if (propertyName == "Shading"):
		retVal = this_Rows.Shading
		try:
			local_ForegroundPatternColorIndex = retVal.ForegroundPatternColorIndex
		except:
			local_ForegroundPatternColorIndex = None
		try:
			local_BackgroundPatternColorIndex = retVal.BackgroundPatternColorIndex
		except:
			local_BackgroundPatternColorIndex = None
		try:
			local_Texture = retVal.Texture
		except:
			local_Texture = None
		try:
			local_ForegroundPatternColor = retVal.ForegroundPatternColor
		except:
			local_ForegroundPatternColor = None
		try:
			local_BackgroundPatternColor = retVal.BackgroundPatternColor
		except:
			local_BackgroundPatternColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Shading", "ForegroundPatternColorIndex": local_ForegroundPatternColorIndex, "BackgroundPatternColorIndex": local_BackgroundPatternColorIndex, "Texture": local_Texture, "ForegroundPatternColor": local_ForegroundPatternColor, "BackgroundPatternColor": local_BackgroundPatternColor, }
	if (propertyName == "WrapAroundText"):
		retVal = this_Rows.WrapAroundText
		return retVal
	if (propertyName == "DistanceTop"):
		retVal = this_Rows.DistanceTop
		return retVal
	if (propertyName == "DistanceBottom"):
		retVal = this_Rows.DistanceBottom
		return retVal
	if (propertyName == "DistanceLeft"):
		retVal = this_Rows.DistanceLeft
		return retVal
	if (propertyName == "DistanceRight"):
		retVal = this_Rows.DistanceRight
		return retVal
	if (propertyName == "HorizontalPosition"):
		retVal = this_Rows.HorizontalPosition
		return retVal
	if (propertyName == "VerticalPosition"):
		retVal = this_Rows.VerticalPosition
		return retVal
	if (propertyName == "RelativeHorizontalPosition"):
		retVal = this_Rows.RelativeHorizontalPosition
		return retVal
	if (propertyName == "RelativeVerticalPosition"):
		retVal = this_Rows.RelativeVerticalPosition
		return retVal
	if (propertyName == "AllowOverlap"):
		retVal = this_Rows.AllowOverlap
		return retVal
	if (propertyName == "NestingLevel"):
		retVal = this_Rows.NestingLevel
		return retVal
	if (propertyName == "TableDirection"):
		retVal = this_Rows.TableDirection
		return retVal


# Tool: 380
@mcp.tool()
async def word_Rows_set_Property(this_Rows_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Rows
	
	propertyName: Name of the property. Can be one of ...
		AllowBreakAcrossPages, Alignment, HeadingFormat, SpaceBetweenColumns, Height, HeightRule, LeftIndent, Borders, WrapAroundText, DistanceTop, DistanceBottom, DistanceLeft, DistanceRight, HorizontalPosition, VerticalPosition, RelativeHorizontalPosition, RelativeVerticalPosition, AllowOverlap, TableDirection
	"""
	this_Rows = get_object(this_Rows_wordObjId)
	
	EnsureWord()
	if (propertyName == "AllowBreakAcrossPages"):
		this_Rows.AllowBreakAcrossPages = propertyValue
	if (propertyName == "Alignment"):
		this_Rows.Alignment = propertyValue
	if (propertyName == "HeadingFormat"):
		this_Rows.HeadingFormat = propertyValue
	if (propertyName == "SpaceBetweenColumns"):
		this_Rows.SpaceBetweenColumns = propertyValue
	if (propertyName == "Height"):
		this_Rows.Height = propertyValue
	if (propertyName == "HeightRule"):
		this_Rows.HeightRule = propertyValue
	if (propertyName == "LeftIndent"):
		this_Rows.LeftIndent = propertyValue
	if (propertyName == "Borders"):
		this_Rows.Borders = propertyValue
	if (propertyName == "WrapAroundText"):
		this_Rows.WrapAroundText = propertyValue
	if (propertyName == "DistanceTop"):
		this_Rows.DistanceTop = propertyValue
	if (propertyName == "DistanceBottom"):
		this_Rows.DistanceBottom = propertyValue
	if (propertyName == "DistanceLeft"):
		this_Rows.DistanceLeft = propertyValue
	if (propertyName == "DistanceRight"):
		this_Rows.DistanceRight = propertyValue
	if (propertyName == "HorizontalPosition"):
		this_Rows.HorizontalPosition = propertyValue
	if (propertyName == "VerticalPosition"):
		this_Rows.VerticalPosition = propertyValue
	if (propertyName == "RelativeHorizontalPosition"):
		this_Rows.RelativeHorizontalPosition = propertyValue
	if (propertyName == "RelativeVerticalPosition"):
		this_Rows.RelativeVerticalPosition = propertyValue
	if (propertyName == "AllowOverlap"):
		this_Rows.AllowOverlap = propertyValue
	if (propertyName == "TableDirection"):
		this_Rows.TableDirection = propertyValue


# Tool: 381
@mcp.tool()
async def word_Columns_Item(this_Columns_wordObjId: str, Index: int):
	"""This tool calls the Item methodon an Columns object. Pass the __WordObjectId of Columns of the object you want to call the method on as the first parameter
	
	Parameters:
		Index: the Index as int
	"""
	this_Columns = get_object(this_Columns_wordObjId)
	retVal = this_Columns.Item(Index)
	try:
		local_Width = retVal.Width
	except:
		local_Width = None
	try:
		local_IsFirst = retVal.IsFirst
	except:
		local_IsFirst = None
	try:
		local_IsLast = retVal.IsLast
	except:
		local_IsLast = None
	try:
		local_Index = retVal.Index
	except:
		local_Index = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_PreferredWidth = retVal.PreferredWidth
	except:
		local_PreferredWidth = None
	try:
		local_PreferredWidthType = retVal.PreferredWidthType
	except:
		local_PreferredWidthType = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Column", "Width": local_Width, "IsFirst": local_IsFirst, "IsLast": local_IsLast, "Index": local_Index, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }


# Tool: 382
@mcp.tool()
async def word_Columns_Add(this_Columns_wordObjId: str, BeforeColumn):
	"""This tool calls the Add methodon an Columns object. Pass the __WordObjectId of Columns of the object you want to call the method on as the first parameter
	
	Parameters:
		BeforeColumn: the BeforeColumn as VT_VARIANT
	"""
	this_Columns = get_object(this_Columns_wordObjId)
	BeforeColumn = tryParseString(BeforeColumn)
	retVal = this_Columns.Add(BeforeColumn)
	try:
		local_Width = retVal.Width
	except:
		local_Width = None
	try:
		local_IsFirst = retVal.IsFirst
	except:
		local_IsFirst = None
	try:
		local_IsLast = retVal.IsLast
	except:
		local_IsLast = None
	try:
		local_Index = retVal.Index
	except:
		local_Index = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_PreferredWidth = retVal.PreferredWidth
	except:
		local_PreferredWidth = None
	try:
		local_PreferredWidthType = retVal.PreferredWidthType
	except:
		local_PreferredWidthType = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Column", "Width": local_Width, "IsFirst": local_IsFirst, "IsLast": local_IsLast, "Index": local_Index, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }


# Tool: 383
@mcp.tool()
async def word_Columns_Select(this_Columns_wordObjId: str):
	"""This tool calls the Select methodon an Columns object. Pass the __WordObjectId of Columns of the object you want to call the method on as the first parameter
"""
	this_Columns = get_object(this_Columns_wordObjId)
	this_Columns.Select()


# Tool: 384
@mcp.tool()
async def word_Columns_Delete(this_Columns_wordObjId: str):
	"""This tool calls the Delete methodon an Columns object. Pass the __WordObjectId of Columns of the object you want to call the method on as the first parameter
"""
	this_Columns = get_object(this_Columns_wordObjId)
	this_Columns.Delete()


# Tool: 385
@mcp.tool()
async def word_Columns_SetWidth(this_Columns_wordObjId: str, ColumnWidth: float, RulerStyle: int):
	"""This tool calls the SetWidth methodon an Columns object. Pass the __WordObjectId of Columns of the object you want to call the method on as the first parameter
	
	Parameters:
		ColumnWidth: the ColumnWidth as float
		RulerStyle: the RulerStyle as WdRulerStyle
	"""
	this_Columns = get_object(this_Columns_wordObjId)
	this_Columns.SetWidth(ColumnWidth, RulerStyle)


# Tool: 386
@mcp.tool()
async def word_Columns_AutoFit(this_Columns_wordObjId: str):
	"""This tool calls the AutoFit methodon an Columns object. Pass the __WordObjectId of Columns of the object you want to call the method on as the first parameter
"""
	this_Columns = get_object(this_Columns_wordObjId)
	this_Columns.AutoFit()


# Tool: 387
@mcp.tool()
async def word_Columns_DistributeWidth(this_Columns_wordObjId: str):
	"""This tool calls the DistributeWidth methodon an Columns object. Pass the __WordObjectId of Columns of the object you want to call the method on as the first parameter
"""
	this_Columns = get_object(this_Columns_wordObjId)
	this_Columns.DistributeWidth()


# Tool: 388
@mcp.tool()
async def word_Columns_get_Property(this_Columns_wordObjId: str, propertyName: str):
	"""Gets properties of Columns
	
	propertyName: Name of the property. Can be one of ...
		Count, First, Last, Width, Borders, Shading, NestingLevel, PreferredWidth, PreferredWidthType
	"""
	this_Columns = get_object(this_Columns_wordObjId)
	
	EnsureWord()
	if (propertyName == "Count"):
		retVal = this_Columns.Count
		return retVal
	if (propertyName == "First"):
		retVal = this_Columns.First
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_IsFirst = retVal.IsFirst
		except:
			local_IsFirst = None
		try:
			local_IsLast = retVal.IsLast
		except:
			local_IsLast = None
		try:
			local_Index = retVal.Index
		except:
			local_Index = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Column", "Width": local_Width, "IsFirst": local_IsFirst, "IsLast": local_IsLast, "Index": local_Index, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "Last"):
		retVal = this_Columns.Last
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_IsFirst = retVal.IsFirst
		except:
			local_IsFirst = None
		try:
			local_IsLast = retVal.IsLast
		except:
			local_IsLast = None
		try:
			local_Index = retVal.Index
		except:
			local_Index = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Column", "Width": local_Width, "IsFirst": local_IsFirst, "IsLast": local_IsLast, "Index": local_Index, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "Width"):
		retVal = this_Columns.Width
		return retVal
	if (propertyName == "Borders"):
		retVal = this_Columns.Borders
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Enable = retVal.Enable
		except:
			local_Enable = None
		try:
			local_DistanceFromTop = retVal.DistanceFromTop
		except:
			local_DistanceFromTop = None
		try:
			local_Shadow = retVal.Shadow
		except:
			local_Shadow = None
		try:
			local_InsideLineStyle = retVal.InsideLineStyle
		except:
			local_InsideLineStyle = None
		try:
			local_OutsideLineStyle = retVal.OutsideLineStyle
		except:
			local_OutsideLineStyle = None
		try:
			local_InsideLineWidth = retVal.InsideLineWidth
		except:
			local_InsideLineWidth = None
		try:
			local_OutsideLineWidth = retVal.OutsideLineWidth
		except:
			local_OutsideLineWidth = None
		try:
			local_InsideColorIndex = retVal.InsideColorIndex
		except:
			local_InsideColorIndex = None
		try:
			local_OutsideColorIndex = retVal.OutsideColorIndex
		except:
			local_OutsideColorIndex = None
		try:
			local_DistanceFromLeft = retVal.DistanceFromLeft
		except:
			local_DistanceFromLeft = None
		try:
			local_DistanceFromBottom = retVal.DistanceFromBottom
		except:
			local_DistanceFromBottom = None
		try:
			local_DistanceFromRight = retVal.DistanceFromRight
		except:
			local_DistanceFromRight = None
		try:
			local_AlwaysInFront = retVal.AlwaysInFront
		except:
			local_AlwaysInFront = None
		try:
			local_SurroundHeader = retVal.SurroundHeader
		except:
			local_SurroundHeader = None
		try:
			local_SurroundFooter = retVal.SurroundFooter
		except:
			local_SurroundFooter = None
		try:
			local_JoinBorders = retVal.JoinBorders
		except:
			local_JoinBorders = None
		try:
			local_HasHorizontal = retVal.HasHorizontal
		except:
			local_HasHorizontal = None
		try:
			local_HasVertical = retVal.HasVertical
		except:
			local_HasVertical = None
		try:
			local_DistanceFrom = retVal.DistanceFrom
		except:
			local_DistanceFrom = None
		try:
			local_EnableFirstPageInSection = retVal.EnableFirstPageInSection
		except:
			local_EnableFirstPageInSection = None
		try:
			local_EnableOtherPagesInSection = retVal.EnableOtherPagesInSection
		except:
			local_EnableOtherPagesInSection = None
		try:
			local_InsideColor = retVal.InsideColor
		except:
			local_InsideColor = None
		try:
			local_OutsideColor = retVal.OutsideColor
		except:
			local_OutsideColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Borders", "Count": local_Count, "Enable": local_Enable, "DistanceFromTop": local_DistanceFromTop, "Shadow": local_Shadow, "InsideLineStyle": local_InsideLineStyle, "OutsideLineStyle": local_OutsideLineStyle, "InsideLineWidth": local_InsideLineWidth, "OutsideLineWidth": local_OutsideLineWidth, "InsideColorIndex": local_InsideColorIndex, "OutsideColorIndex": local_OutsideColorIndex, "DistanceFromLeft": local_DistanceFromLeft, "DistanceFromBottom": local_DistanceFromBottom, "DistanceFromRight": local_DistanceFromRight, "AlwaysInFront": local_AlwaysInFront, "SurroundHeader": local_SurroundHeader, "SurroundFooter": local_SurroundFooter, "JoinBorders": local_JoinBorders, "HasHorizontal": local_HasHorizontal, "HasVertical": local_HasVertical, "DistanceFrom": local_DistanceFrom, "EnableFirstPageInSection": local_EnableFirstPageInSection, "EnableOtherPagesInSection": local_EnableOtherPagesInSection, "InsideColor": local_InsideColor, "OutsideColor": local_OutsideColor, }
	if (propertyName == "Shading"):
		retVal = this_Columns.Shading
		try:
			local_ForegroundPatternColorIndex = retVal.ForegroundPatternColorIndex
		except:
			local_ForegroundPatternColorIndex = None
		try:
			local_BackgroundPatternColorIndex = retVal.BackgroundPatternColorIndex
		except:
			local_BackgroundPatternColorIndex = None
		try:
			local_Texture = retVal.Texture
		except:
			local_Texture = None
		try:
			local_ForegroundPatternColor = retVal.ForegroundPatternColor
		except:
			local_ForegroundPatternColor = None
		try:
			local_BackgroundPatternColor = retVal.BackgroundPatternColor
		except:
			local_BackgroundPatternColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Shading", "ForegroundPatternColorIndex": local_ForegroundPatternColorIndex, "BackgroundPatternColorIndex": local_BackgroundPatternColorIndex, "Texture": local_Texture, "ForegroundPatternColor": local_ForegroundPatternColor, "BackgroundPatternColor": local_BackgroundPatternColor, }
	if (propertyName == "NestingLevel"):
		retVal = this_Columns.NestingLevel
		return retVal
	if (propertyName == "PreferredWidth"):
		retVal = this_Columns.PreferredWidth
		return retVal
	if (propertyName == "PreferredWidthType"):
		retVal = this_Columns.PreferredWidthType
		return retVal


# Tool: 389
@mcp.tool()
async def word_Columns_set_Property(this_Columns_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Columns
	
	propertyName: Name of the property. Can be one of ...
		Width, Borders, PreferredWidth, PreferredWidthType
	"""
	this_Columns = get_object(this_Columns_wordObjId)
	
	EnsureWord()
	if (propertyName == "Width"):
		this_Columns.Width = propertyValue
	if (propertyName == "Borders"):
		this_Columns.Borders = propertyValue
	if (propertyName == "PreferredWidth"):
		this_Columns.PreferredWidth = propertyValue
	if (propertyName == "PreferredWidthType"):
		this_Columns.PreferredWidthType = propertyValue


# Tool: 390
@mcp.tool()
async def word_Cells_Item(this_Cells_wordObjId: str, Index: int):
	"""This tool calls the Item methodon an Cells object. Pass the __WordObjectId of Cells of the object you want to call the method on as the first parameter
	
	Parameters:
		Index: the Index as int
	"""
	this_Cells = get_object(this_Cells_wordObjId)
	retVal = this_Cells.Item(Index)
	try:
		local_RowIndex = retVal.RowIndex
	except:
		local_RowIndex = None
	try:
		local_ColumnIndex = retVal.ColumnIndex
	except:
		local_ColumnIndex = None
	try:
		local_Width = retVal.Width
	except:
		local_Width = None
	try:
		local_Height = retVal.Height
	except:
		local_Height = None
	try:
		local_HeightRule = retVal.HeightRule
	except:
		local_HeightRule = None
	try:
		local_VerticalAlignment = retVal.VerticalAlignment
	except:
		local_VerticalAlignment = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_WordWrap = retVal.WordWrap
	except:
		local_WordWrap = None
	try:
		local_PreferredWidth = retVal.PreferredWidth
	except:
		local_PreferredWidth = None
	try:
		local_FitText = retVal.FitText
	except:
		local_FitText = None
	try:
		local_TopPadding = retVal.TopPadding
	except:
		local_TopPadding = None
	try:
		local_BottomPadding = retVal.BottomPadding
	except:
		local_BottomPadding = None
	try:
		local_LeftPadding = retVal.LeftPadding
	except:
		local_LeftPadding = None
	try:
		local_RightPadding = retVal.RightPadding
	except:
		local_RightPadding = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_PreferredWidthType = retVal.PreferredWidthType
	except:
		local_PreferredWidthType = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Cell", "RowIndex": local_RowIndex, "ColumnIndex": local_ColumnIndex, "Width": local_Width, "Height": local_Height, "HeightRule": local_HeightRule, "VerticalAlignment": local_VerticalAlignment, "NestingLevel": local_NestingLevel, "WordWrap": local_WordWrap, "PreferredWidth": local_PreferredWidth, "FitText": local_FitText, "TopPadding": local_TopPadding, "BottomPadding": local_BottomPadding, "LeftPadding": local_LeftPadding, "RightPadding": local_RightPadding, "ID": local_ID, "PreferredWidthType": local_PreferredWidthType, }


# Tool: 391
@mcp.tool()
async def word_Cells_Add(this_Cells_wordObjId: str, BeforeCell):
	"""This tool calls the Add methodon an Cells object. Pass the __WordObjectId of Cells of the object you want to call the method on as the first parameter
	
	Parameters:
		BeforeCell: the BeforeCell as VT_VARIANT
	"""
	this_Cells = get_object(this_Cells_wordObjId)
	BeforeCell = tryParseString(BeforeCell)
	retVal = this_Cells.Add(BeforeCell)
	try:
		local_RowIndex = retVal.RowIndex
	except:
		local_RowIndex = None
	try:
		local_ColumnIndex = retVal.ColumnIndex
	except:
		local_ColumnIndex = None
	try:
		local_Width = retVal.Width
	except:
		local_Width = None
	try:
		local_Height = retVal.Height
	except:
		local_Height = None
	try:
		local_HeightRule = retVal.HeightRule
	except:
		local_HeightRule = None
	try:
		local_VerticalAlignment = retVal.VerticalAlignment
	except:
		local_VerticalAlignment = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_WordWrap = retVal.WordWrap
	except:
		local_WordWrap = None
	try:
		local_PreferredWidth = retVal.PreferredWidth
	except:
		local_PreferredWidth = None
	try:
		local_FitText = retVal.FitText
	except:
		local_FitText = None
	try:
		local_TopPadding = retVal.TopPadding
	except:
		local_TopPadding = None
	try:
		local_BottomPadding = retVal.BottomPadding
	except:
		local_BottomPadding = None
	try:
		local_LeftPadding = retVal.LeftPadding
	except:
		local_LeftPadding = None
	try:
		local_RightPadding = retVal.RightPadding
	except:
		local_RightPadding = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_PreferredWidthType = retVal.PreferredWidthType
	except:
		local_PreferredWidthType = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Cell", "RowIndex": local_RowIndex, "ColumnIndex": local_ColumnIndex, "Width": local_Width, "Height": local_Height, "HeightRule": local_HeightRule, "VerticalAlignment": local_VerticalAlignment, "NestingLevel": local_NestingLevel, "WordWrap": local_WordWrap, "PreferredWidth": local_PreferredWidth, "FitText": local_FitText, "TopPadding": local_TopPadding, "BottomPadding": local_BottomPadding, "LeftPadding": local_LeftPadding, "RightPadding": local_RightPadding, "ID": local_ID, "PreferredWidthType": local_PreferredWidthType, }


# Tool: 392
@mcp.tool()
async def word_Cells_Delete(this_Cells_wordObjId: str, ShiftCells):
	"""This tool calls the Delete methodon an Cells object. Pass the __WordObjectId of Cells of the object you want to call the method on as the first parameter
	
	Parameters:
		ShiftCells: the ShiftCells as VT_VARIANT
	"""
	this_Cells = get_object(this_Cells_wordObjId)
	ShiftCells = tryParseString(ShiftCells)
	this_Cells.Delete(ShiftCells)


# Tool: 393
@mcp.tool()
async def word_Cells_SetWidth(this_Cells_wordObjId: str, ColumnWidth: float, RulerStyle: int):
	"""This tool calls the SetWidth methodon an Cells object. Pass the __WordObjectId of Cells of the object you want to call the method on as the first parameter
	
	Parameters:
		ColumnWidth: the ColumnWidth as float
		RulerStyle: the RulerStyle as WdRulerStyle
	"""
	this_Cells = get_object(this_Cells_wordObjId)
	this_Cells.SetWidth(ColumnWidth, RulerStyle)


# Tool: 394
@mcp.tool()
async def word_Cells_SetHeight(this_Cells_wordObjId: str, RowHeight, HeightRule: int):
	"""This tool calls the SetHeight methodon an Cells object. Pass the __WordObjectId of Cells of the object you want to call the method on as the first parameter
	
	Parameters:
		RowHeight: the RowHeight as VT_VARIANT
		HeightRule: the HeightRule as WdRowHeightRule
	"""
	this_Cells = get_object(this_Cells_wordObjId)
	RowHeight = tryParseString(RowHeight)
	this_Cells.SetHeight(RowHeight, HeightRule)


# Tool: 395
@mcp.tool()
async def word_Cells_Merge(this_Cells_wordObjId: str):
	"""This tool calls the Merge methodon an Cells object. Pass the __WordObjectId of Cells of the object you want to call the method on as the first parameter
"""
	this_Cells = get_object(this_Cells_wordObjId)
	this_Cells.Merge()


# Tool: 396
@mcp.tool()
async def word_Cells_Split(this_Cells_wordObjId: str, NumRows, NumColumns, MergeBeforeSplit):
	"""This tool calls the Split methodon an Cells object. Pass the __WordObjectId of Cells of the object you want to call the method on as the first parameter
	
	Parameters:
		NumRows: the NumRows as VT_VARIANT
		NumColumns: the NumColumns as VT_VARIANT
		MergeBeforeSplit: the MergeBeforeSplit as VT_VARIANT
	"""
	this_Cells = get_object(this_Cells_wordObjId)
	NumRows = tryParseString(NumRows)
	NumColumns = tryParseString(NumColumns)
	MergeBeforeSplit = tryParseString(MergeBeforeSplit)
	this_Cells.Split(NumRows, NumColumns, MergeBeforeSplit)


# Tool: 397
@mcp.tool()
async def word_Cells_DistributeHeight(this_Cells_wordObjId: str):
	"""This tool calls the DistributeHeight methodon an Cells object. Pass the __WordObjectId of Cells of the object you want to call the method on as the first parameter
"""
	this_Cells = get_object(this_Cells_wordObjId)
	this_Cells.DistributeHeight()


# Tool: 398
@mcp.tool()
async def word_Cells_DistributeWidth(this_Cells_wordObjId: str):
	"""This tool calls the DistributeWidth methodon an Cells object. Pass the __WordObjectId of Cells of the object you want to call the method on as the first parameter
"""
	this_Cells = get_object(this_Cells_wordObjId)
	this_Cells.DistributeWidth()


# Tool: 399
@mcp.tool()
async def word_Cells_AutoFit(this_Cells_wordObjId: str):
	"""This tool calls the AutoFit methodon an Cells object. Pass the __WordObjectId of Cells of the object you want to call the method on as the first parameter
"""
	this_Cells = get_object(this_Cells_wordObjId)
	this_Cells.AutoFit()


# Tool: 400
@mcp.tool()
async def word_Cells_get_Property(this_Cells_wordObjId: str, propertyName: str):
	"""Gets properties of Cells
	
	propertyName: Name of the property. Can be one of ...
		Count, Width, Height, HeightRule, VerticalAlignment, Borders, Shading, NestingLevel, PreferredWidth, PreferredWidthType
	"""
	this_Cells = get_object(this_Cells_wordObjId)
	
	EnsureWord()
	if (propertyName == "Count"):
		retVal = this_Cells.Count
		return retVal
	if (propertyName == "Width"):
		retVal = this_Cells.Width
		return retVal
	if (propertyName == "Height"):
		retVal = this_Cells.Height
		return retVal
	if (propertyName == "HeightRule"):
		retVal = this_Cells.HeightRule
		return retVal
	if (propertyName == "VerticalAlignment"):
		retVal = this_Cells.VerticalAlignment
		return retVal
	if (propertyName == "Borders"):
		retVal = this_Cells.Borders
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Enable = retVal.Enable
		except:
			local_Enable = None
		try:
			local_DistanceFromTop = retVal.DistanceFromTop
		except:
			local_DistanceFromTop = None
		try:
			local_Shadow = retVal.Shadow
		except:
			local_Shadow = None
		try:
			local_InsideLineStyle = retVal.InsideLineStyle
		except:
			local_InsideLineStyle = None
		try:
			local_OutsideLineStyle = retVal.OutsideLineStyle
		except:
			local_OutsideLineStyle = None
		try:
			local_InsideLineWidth = retVal.InsideLineWidth
		except:
			local_InsideLineWidth = None
		try:
			local_OutsideLineWidth = retVal.OutsideLineWidth
		except:
			local_OutsideLineWidth = None
		try:
			local_InsideColorIndex = retVal.InsideColorIndex
		except:
			local_InsideColorIndex = None
		try:
			local_OutsideColorIndex = retVal.OutsideColorIndex
		except:
			local_OutsideColorIndex = None
		try:
			local_DistanceFromLeft = retVal.DistanceFromLeft
		except:
			local_DistanceFromLeft = None
		try:
			local_DistanceFromBottom = retVal.DistanceFromBottom
		except:
			local_DistanceFromBottom = None
		try:
			local_DistanceFromRight = retVal.DistanceFromRight
		except:
			local_DistanceFromRight = None
		try:
			local_AlwaysInFront = retVal.AlwaysInFront
		except:
			local_AlwaysInFront = None
		try:
			local_SurroundHeader = retVal.SurroundHeader
		except:
			local_SurroundHeader = None
		try:
			local_SurroundFooter = retVal.SurroundFooter
		except:
			local_SurroundFooter = None
		try:
			local_JoinBorders = retVal.JoinBorders
		except:
			local_JoinBorders = None
		try:
			local_HasHorizontal = retVal.HasHorizontal
		except:
			local_HasHorizontal = None
		try:
			local_HasVertical = retVal.HasVertical
		except:
			local_HasVertical = None
		try:
			local_DistanceFrom = retVal.DistanceFrom
		except:
			local_DistanceFrom = None
		try:
			local_EnableFirstPageInSection = retVal.EnableFirstPageInSection
		except:
			local_EnableFirstPageInSection = None
		try:
			local_EnableOtherPagesInSection = retVal.EnableOtherPagesInSection
		except:
			local_EnableOtherPagesInSection = None
		try:
			local_InsideColor = retVal.InsideColor
		except:
			local_InsideColor = None
		try:
			local_OutsideColor = retVal.OutsideColor
		except:
			local_OutsideColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Borders", "Count": local_Count, "Enable": local_Enable, "DistanceFromTop": local_DistanceFromTop, "Shadow": local_Shadow, "InsideLineStyle": local_InsideLineStyle, "OutsideLineStyle": local_OutsideLineStyle, "InsideLineWidth": local_InsideLineWidth, "OutsideLineWidth": local_OutsideLineWidth, "InsideColorIndex": local_InsideColorIndex, "OutsideColorIndex": local_OutsideColorIndex, "DistanceFromLeft": local_DistanceFromLeft, "DistanceFromBottom": local_DistanceFromBottom, "DistanceFromRight": local_DistanceFromRight, "AlwaysInFront": local_AlwaysInFront, "SurroundHeader": local_SurroundHeader, "SurroundFooter": local_SurroundFooter, "JoinBorders": local_JoinBorders, "HasHorizontal": local_HasHorizontal, "HasVertical": local_HasVertical, "DistanceFrom": local_DistanceFrom, "EnableFirstPageInSection": local_EnableFirstPageInSection, "EnableOtherPagesInSection": local_EnableOtherPagesInSection, "InsideColor": local_InsideColor, "OutsideColor": local_OutsideColor, }
	if (propertyName == "Shading"):
		retVal = this_Cells.Shading
		try:
			local_ForegroundPatternColorIndex = retVal.ForegroundPatternColorIndex
		except:
			local_ForegroundPatternColorIndex = None
		try:
			local_BackgroundPatternColorIndex = retVal.BackgroundPatternColorIndex
		except:
			local_BackgroundPatternColorIndex = None
		try:
			local_Texture = retVal.Texture
		except:
			local_Texture = None
		try:
			local_ForegroundPatternColor = retVal.ForegroundPatternColor
		except:
			local_ForegroundPatternColor = None
		try:
			local_BackgroundPatternColor = retVal.BackgroundPatternColor
		except:
			local_BackgroundPatternColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Shading", "ForegroundPatternColorIndex": local_ForegroundPatternColorIndex, "BackgroundPatternColorIndex": local_BackgroundPatternColorIndex, "Texture": local_Texture, "ForegroundPatternColor": local_ForegroundPatternColor, "BackgroundPatternColor": local_BackgroundPatternColor, }
	if (propertyName == "NestingLevel"):
		retVal = this_Cells.NestingLevel
		return retVal
	if (propertyName == "PreferredWidth"):
		retVal = this_Cells.PreferredWidth
		return retVal
	if (propertyName == "PreferredWidthType"):
		retVal = this_Cells.PreferredWidthType
		return retVal


# Tool: 401
@mcp.tool()
async def word_Cells_set_Property(this_Cells_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Cells
	
	propertyName: Name of the property. Can be one of ...
		Width, Height, HeightRule, VerticalAlignment, Borders, PreferredWidth, PreferredWidthType
	"""
	this_Cells = get_object(this_Cells_wordObjId)
	
	EnsureWord()
	if (propertyName == "Width"):
		this_Cells.Width = propertyValue
	if (propertyName == "Height"):
		this_Cells.Height = propertyValue
	if (propertyName == "HeightRule"):
		this_Cells.HeightRule = propertyValue
	if (propertyName == "VerticalAlignment"):
		this_Cells.VerticalAlignment = propertyValue
	if (propertyName == "Borders"):
		this_Cells.Borders = propertyValue
	if (propertyName == "PreferredWidth"):
		this_Cells.PreferredWidth = propertyValue
	if (propertyName == "PreferredWidthType"):
		this_Cells.PreferredWidthType = propertyValue


# Tool: 402
@mcp.tool()
async def word_Borders_Item(this_Borders_wordObjId: str, Index: int):
	"""This tool calls the Item methodon an Borders object. Pass the __WordObjectId of Borders of the object you want to call the method on as the first parameter
	
	Parameters:
		Index: the Index as WdBorderType
	"""
	this_Borders = get_object(this_Borders_wordObjId)
	retVal = this_Borders.Item(Index)
	try:
		local_Visible = retVal.Visible
	except:
		local_Visible = None
	try:
		local_ColorIndex = retVal.ColorIndex
	except:
		local_ColorIndex = None
	try:
		local_Inside = retVal.Inside
	except:
		local_Inside = None
	try:
		local_LineStyle = retVal.LineStyle
	except:
		local_LineStyle = None
	try:
		local_LineWidth = retVal.LineWidth
	except:
		local_LineWidth = None
	try:
		local_ArtStyle = retVal.ArtStyle
	except:
		local_ArtStyle = None
	try:
		local_ArtWidth = retVal.ArtWidth
	except:
		local_ArtWidth = None
	try:
		local_Color = retVal.Color
	except:
		local_Color = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Border", "Visible": local_Visible, "ColorIndex": local_ColorIndex, "Inside": local_Inside, "LineStyle": local_LineStyle, "LineWidth": local_LineWidth, "ArtStyle": local_ArtStyle, "ArtWidth": local_ArtWidth, "Color": local_Color, }


# Tool: 403
@mcp.tool()
async def word_Borders_ApplyPageBordersToAllSections(this_Borders_wordObjId: str):
	"""This tool calls the ApplyPageBordersToAllSections methodon an Borders object. Pass the __WordObjectId of Borders of the object you want to call the method on as the first parameter
"""
	this_Borders = get_object(this_Borders_wordObjId)
	this_Borders.ApplyPageBordersToAllSections()


# Tool: 404
@mcp.tool()
async def word_Borders_get_Property(this_Borders_wordObjId: str, propertyName: str):
	"""Gets properties of Borders
	
	propertyName: Name of the property. Can be one of ...
		Count, Enable, DistanceFromTop, Shadow, InsideLineStyle, OutsideLineStyle, InsideLineWidth, OutsideLineWidth, InsideColorIndex, OutsideColorIndex, DistanceFromLeft, DistanceFromBottom, DistanceFromRight, AlwaysInFront, SurroundHeader, SurroundFooter, JoinBorders, HasHorizontal, HasVertical, DistanceFrom, EnableFirstPageInSection, EnableOtherPagesInSection, InsideColor, OutsideColor
	"""
	this_Borders = get_object(this_Borders_wordObjId)
	
	EnsureWord()
	if (propertyName == "Count"):
		retVal = this_Borders.Count
		return retVal
	if (propertyName == "Enable"):
		retVal = this_Borders.Enable
		return retVal
	if (propertyName == "DistanceFromTop"):
		retVal = this_Borders.DistanceFromTop
		return retVal
	if (propertyName == "Shadow"):
		retVal = this_Borders.Shadow
		return retVal
	if (propertyName == "InsideLineStyle"):
		retVal = this_Borders.InsideLineStyle
		return retVal
	if (propertyName == "OutsideLineStyle"):
		retVal = this_Borders.OutsideLineStyle
		return retVal
	if (propertyName == "InsideLineWidth"):
		retVal = this_Borders.InsideLineWidth
		return retVal
	if (propertyName == "OutsideLineWidth"):
		retVal = this_Borders.OutsideLineWidth
		return retVal
	if (propertyName == "InsideColorIndex"):
		retVal = this_Borders.InsideColorIndex
		return retVal
	if (propertyName == "OutsideColorIndex"):
		retVal = this_Borders.OutsideColorIndex
		return retVal
	if (propertyName == "DistanceFromLeft"):
		retVal = this_Borders.DistanceFromLeft
		return retVal
	if (propertyName == "DistanceFromBottom"):
		retVal = this_Borders.DistanceFromBottom
		return retVal
	if (propertyName == "DistanceFromRight"):
		retVal = this_Borders.DistanceFromRight
		return retVal
	if (propertyName == "AlwaysInFront"):
		retVal = this_Borders.AlwaysInFront
		return retVal
	if (propertyName == "SurroundHeader"):
		retVal = this_Borders.SurroundHeader
		return retVal
	if (propertyName == "SurroundFooter"):
		retVal = this_Borders.SurroundFooter
		return retVal
	if (propertyName == "JoinBorders"):
		retVal = this_Borders.JoinBorders
		return retVal
	if (propertyName == "HasHorizontal"):
		retVal = this_Borders.HasHorizontal
		return retVal
	if (propertyName == "HasVertical"):
		retVal = this_Borders.HasVertical
		return retVal
	if (propertyName == "DistanceFrom"):
		retVal = this_Borders.DistanceFrom
		return retVal
	if (propertyName == "EnableFirstPageInSection"):
		retVal = this_Borders.EnableFirstPageInSection
		return retVal
	if (propertyName == "EnableOtherPagesInSection"):
		retVal = this_Borders.EnableOtherPagesInSection
		return retVal
	if (propertyName == "InsideColor"):
		retVal = this_Borders.InsideColor
		return retVal
	if (propertyName == "OutsideColor"):
		retVal = this_Borders.OutsideColor
		return retVal


# Tool: 405
@mcp.tool()
async def word_Borders_set_Property(this_Borders_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Borders
	
	propertyName: Name of the property. Can be one of ...
		Enable, DistanceFromTop, Shadow, InsideLineStyle, OutsideLineStyle, InsideLineWidth, OutsideLineWidth, InsideColorIndex, OutsideColorIndex, DistanceFromLeft, DistanceFromBottom, DistanceFromRight, AlwaysInFront, SurroundHeader, SurroundFooter, JoinBorders, DistanceFrom, EnableFirstPageInSection, EnableOtherPagesInSection, InsideColor, OutsideColor
	"""
	this_Borders = get_object(this_Borders_wordObjId)
	
	EnsureWord()
	if (propertyName == "Enable"):
		this_Borders.Enable = propertyValue
	if (propertyName == "DistanceFromTop"):
		this_Borders.DistanceFromTop = propertyValue
	if (propertyName == "Shadow"):
		this_Borders.Shadow = propertyValue
	if (propertyName == "InsideLineStyle"):
		this_Borders.InsideLineStyle = propertyValue
	if (propertyName == "OutsideLineStyle"):
		this_Borders.OutsideLineStyle = propertyValue
	if (propertyName == "InsideLineWidth"):
		this_Borders.InsideLineWidth = propertyValue
	if (propertyName == "OutsideLineWidth"):
		this_Borders.OutsideLineWidth = propertyValue
	if (propertyName == "InsideColorIndex"):
		this_Borders.InsideColorIndex = propertyValue
	if (propertyName == "OutsideColorIndex"):
		this_Borders.OutsideColorIndex = propertyValue
	if (propertyName == "DistanceFromLeft"):
		this_Borders.DistanceFromLeft = propertyValue
	if (propertyName == "DistanceFromBottom"):
		this_Borders.DistanceFromBottom = propertyValue
	if (propertyName == "DistanceFromRight"):
		this_Borders.DistanceFromRight = propertyValue
	if (propertyName == "AlwaysInFront"):
		this_Borders.AlwaysInFront = propertyValue
	if (propertyName == "SurroundHeader"):
		this_Borders.SurroundHeader = propertyValue
	if (propertyName == "SurroundFooter"):
		this_Borders.SurroundFooter = propertyValue
	if (propertyName == "JoinBorders"):
		this_Borders.JoinBorders = propertyValue
	if (propertyName == "DistanceFrom"):
		this_Borders.DistanceFrom = propertyValue
	if (propertyName == "EnableFirstPageInSection"):
		this_Borders.EnableFirstPageInSection = propertyValue
	if (propertyName == "EnableOtherPagesInSection"):
		this_Borders.EnableOtherPagesInSection = propertyValue
	if (propertyName == "InsideColor"):
		this_Borders.InsideColor = propertyValue
	if (propertyName == "OutsideColor"):
		this_Borders.OutsideColor = propertyValue


# Tool: 406
@mcp.tool()
async def word_Border_get_Property(this_Border_wordObjId: str, propertyName: str):
	"""Gets properties of Border
	
	propertyName: Name of the property. Can be one of ...
		Visible, ColorIndex, Inside, LineStyle, LineWidth, ArtStyle, ArtWidth, Color
	"""
	this_Border = get_object(this_Border_wordObjId)
	
	EnsureWord()
	if (propertyName == "Visible"):
		retVal = this_Border.Visible
		return retVal
	if (propertyName == "ColorIndex"):
		retVal = this_Border.ColorIndex
		return retVal
	if (propertyName == "Inside"):
		retVal = this_Border.Inside
		return retVal
	if (propertyName == "LineStyle"):
		retVal = this_Border.LineStyle
		return retVal
	if (propertyName == "LineWidth"):
		retVal = this_Border.LineWidth
		return retVal
	if (propertyName == "ArtStyle"):
		retVal = this_Border.ArtStyle
		return retVal
	if (propertyName == "ArtWidth"):
		retVal = this_Border.ArtWidth
		return retVal
	if (propertyName == "Color"):
		retVal = this_Border.Color
		return retVal


# Tool: 407
@mcp.tool()
async def word_Border_set_Property(this_Border_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Border
	
	propertyName: Name of the property. Can be one of ...
		Visible, ColorIndex, LineStyle, LineWidth, ArtStyle, ArtWidth, Color
	"""
	this_Border = get_object(this_Border_wordObjId)
	
	EnsureWord()
	if (propertyName == "Visible"):
		this_Border.Visible = propertyValue
	if (propertyName == "ColorIndex"):
		this_Border.ColorIndex = propertyValue
	if (propertyName == "LineStyle"):
		this_Border.LineStyle = propertyValue
	if (propertyName == "LineWidth"):
		this_Border.LineWidth = propertyValue
	if (propertyName == "ArtStyle"):
		this_Border.ArtStyle = propertyValue
	if (propertyName == "ArtWidth"):
		this_Border.ArtWidth = propertyValue
	if (propertyName == "Color"):
		this_Border.Color = propertyValue


# Tool: 408
@mcp.tool()
async def word_Shading_get_Property(this_Shading_wordObjId: str, propertyName: str):
	"""Gets properties of Shading
	
	propertyName: Name of the property. Can be one of ...
		ForegroundPatternColorIndex, BackgroundPatternColorIndex, Texture, ForegroundPatternColor, BackgroundPatternColor
	"""
	this_Shading = get_object(this_Shading_wordObjId)
	
	EnsureWord()
	if (propertyName == "ForegroundPatternColorIndex"):
		retVal = this_Shading.ForegroundPatternColorIndex
		return retVal
	if (propertyName == "BackgroundPatternColorIndex"):
		retVal = this_Shading.BackgroundPatternColorIndex
		return retVal
	if (propertyName == "Texture"):
		retVal = this_Shading.Texture
		return retVal
	if (propertyName == "ForegroundPatternColor"):
		retVal = this_Shading.ForegroundPatternColor
		return retVal
	if (propertyName == "BackgroundPatternColor"):
		retVal = this_Shading.BackgroundPatternColor
		return retVal


# Tool: 409
@mcp.tool()
async def word_Shading_set_Property(this_Shading_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Shading
	
	propertyName: Name of the property. Can be one of ...
		ForegroundPatternColorIndex, BackgroundPatternColorIndex, Texture, ForegroundPatternColor, BackgroundPatternColor
	"""
	this_Shading = get_object(this_Shading_wordObjId)
	
	EnsureWord()
	if (propertyName == "ForegroundPatternColorIndex"):
		this_Shading.ForegroundPatternColorIndex = propertyValue
	if (propertyName == "BackgroundPatternColorIndex"):
		this_Shading.BackgroundPatternColorIndex = propertyValue
	if (propertyName == "Texture"):
		this_Shading.Texture = propertyValue
	if (propertyName == "ForegroundPatternColor"):
		this_Shading.ForegroundPatternColor = propertyValue
	if (propertyName == "BackgroundPatternColor"):
		this_Shading.BackgroundPatternColor = propertyValue


# Tool: 410
@mcp.tool()
async def word_Selection_Select(this_Selection_wordObjId: str):
	"""This tool calls the Select methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.Select()


# Tool: 411
@mcp.tool()
async def word_Selection_SetRange(this_Selection_wordObjId: str, Start: int, End: int):
	"""This tool calls the SetRange methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Start: the Start as int
		End: the End as int
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.SetRange(Start, End)


# Tool: 412
@mcp.tool()
async def word_Selection_Collapse(this_Selection_wordObjId: str, Direction):
	"""This tool calls the Collapse methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Direction: the Direction as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Direction = tryParseString(Direction)
	this_Selection.Collapse(Direction)


# Tool: 413
@mcp.tool()
async def word_Selection_InsertBefore(this_Selection_wordObjId: str, Text: str):
	"""This tool calls the InsertBefore methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Text: the Text as str
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.InsertBefore(Text)


# Tool: 414
@mcp.tool()
async def word_Selection_InsertAfter(this_Selection_wordObjId: str, Text: str):
	"""This tool calls the InsertAfter methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Text: the Text as str
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.InsertAfter(Text)


# Tool: 415
@mcp.tool()
async def word_Selection_Next(this_Selection_wordObjId: str, Unit, Count):
	"""This tool calls the Next methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	retVal = this_Selection.Next(Unit, Count)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 416
@mcp.tool()
async def word_Selection_Previous(this_Selection_wordObjId: str, Unit, Count):
	"""This tool calls the Previous methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	retVal = this_Selection.Previous(Unit, Count)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 417
@mcp.tool()
async def word_Selection_StartOf(this_Selection_wordObjId: str, Unit, Extend):
	"""This tool calls the StartOf methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Extend: the Extend as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Extend = tryParseString(Extend)
	retVal = this_Selection.StartOf(Unit, Extend)
	return retVal


# Tool: 418
@mcp.tool()
async def word_Selection_EndOf(this_Selection_wordObjId: str, Unit, Extend):
	"""This tool calls the EndOf methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Extend: the Extend as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Extend = tryParseString(Extend)
	retVal = this_Selection.EndOf(Unit, Extend)
	return retVal


# Tool: 419
@mcp.tool()
async def word_Selection_Move(this_Selection_wordObjId: str, Unit, Count):
	"""This tool calls the Move methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	retVal = this_Selection.Move(Unit, Count)
	return retVal


# Tool: 420
@mcp.tool()
async def word_Selection_MoveStart(this_Selection_wordObjId: str, Unit, Count):
	"""This tool calls the MoveStart methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	retVal = this_Selection.MoveStart(Unit, Count)
	return retVal


# Tool: 421
@mcp.tool()
async def word_Selection_MoveEnd(this_Selection_wordObjId: str, Unit, Count):
	"""This tool calls the MoveEnd methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	retVal = this_Selection.MoveEnd(Unit, Count)
	return retVal


# Tool: 422
@mcp.tool()
async def word_Selection_MoveWhile(this_Selection_wordObjId: str, Cset, Count):
	"""This tool calls the MoveWhile methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Cset: the Cset as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Cset = tryParseString(Cset)
	Count = tryParseString(Count)
	retVal = this_Selection.MoveWhile(Cset, Count)
	return retVal


# Tool: 423
@mcp.tool()
async def word_Selection_MoveStartWhile(this_Selection_wordObjId: str, Cset, Count):
	"""This tool calls the MoveStartWhile methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Cset: the Cset as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Cset = tryParseString(Cset)
	Count = tryParseString(Count)
	retVal = this_Selection.MoveStartWhile(Cset, Count)
	return retVal


# Tool: 424
@mcp.tool()
async def word_Selection_MoveEndWhile(this_Selection_wordObjId: str, Cset, Count):
	"""This tool calls the MoveEndWhile methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Cset: the Cset as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Cset = tryParseString(Cset)
	Count = tryParseString(Count)
	retVal = this_Selection.MoveEndWhile(Cset, Count)
	return retVal


# Tool: 425
@mcp.tool()
async def word_Selection_MoveUntil(this_Selection_wordObjId: str, Cset, Count):
	"""This tool calls the MoveUntil methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Cset: the Cset as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Cset = tryParseString(Cset)
	Count = tryParseString(Count)
	retVal = this_Selection.MoveUntil(Cset, Count)
	return retVal


# Tool: 426
@mcp.tool()
async def word_Selection_MoveStartUntil(this_Selection_wordObjId: str, Cset, Count):
	"""This tool calls the MoveStartUntil methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Cset: the Cset as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Cset = tryParseString(Cset)
	Count = tryParseString(Count)
	retVal = this_Selection.MoveStartUntil(Cset, Count)
	return retVal


# Tool: 427
@mcp.tool()
async def word_Selection_MoveEndUntil(this_Selection_wordObjId: str, Cset, Count):
	"""This tool calls the MoveEndUntil methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Cset: the Cset as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Cset = tryParseString(Cset)
	Count = tryParseString(Count)
	retVal = this_Selection.MoveEndUntil(Cset, Count)
	return retVal


# Tool: 428
@mcp.tool()
async def word_Selection_Cut(this_Selection_wordObjId: str):
	"""This tool calls the Cut methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.Cut()


# Tool: 429
@mcp.tool()
async def word_Selection_Copy(this_Selection_wordObjId: str):
	"""This tool calls the Copy methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.Copy()


# Tool: 430
@mcp.tool()
async def word_Selection_Paste(this_Selection_wordObjId: str):
	"""This tool calls the Paste methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.Paste()


# Tool: 431
@mcp.tool()
async def word_Selection_InsertBreak(this_Selection_wordObjId: str, Type):
	"""This tool calls the InsertBreak methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Type: the Type as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Type = tryParseString(Type)
	this_Selection.InsertBreak(Type)


# Tool: 432
@mcp.tool()
async def word_Selection_InsertFile(this_Selection_wordObjId: str, FileName: str, Range, ConfirmConversions, Link, Attachment):
	"""This tool calls the InsertFile methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		FileName: the FileName as str
		Range: the Range as VT_VARIANT
		ConfirmConversions: the ConfirmConversions as VT_VARIANT
		Link: the Link as VT_VARIANT
		Attachment: the Attachment as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Range = tryParseString(Range)
	ConfirmConversions = tryParseString(ConfirmConversions)
	Link = tryParseString(Link)
	Attachment = tryParseString(Attachment)
	this_Selection.InsertFile(FileName, Range, ConfirmConversions, Link, Attachment)


# Tool: 433
@mcp.tool()
async def word_Selection_InStory(this_Selection_wordObjId: str, Range_wordObjId: str):
	"""This tool calls the InStory methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Range_wordObjId: 		To pass this object, send in the __WordObjectId of the Range object as was obtained from a previous return value
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Range = get_object(Range_wordObjId)
	retVal = this_Selection.InStory(Range)
	return retVal


# Tool: 434
@mcp.tool()
async def word_Selection_InRange(this_Selection_wordObjId: str, Range_wordObjId: str):
	"""This tool calls the InRange methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Range_wordObjId: 		To pass this object, send in the __WordObjectId of the Range object as was obtained from a previous return value
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Range = get_object(Range_wordObjId)
	retVal = this_Selection.InRange(Range)
	return retVal


# Tool: 435
@mcp.tool()
async def word_Selection_Delete(this_Selection_wordObjId: str, Unit, Count):
	"""This tool calls the Delete methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	retVal = this_Selection.Delete(Unit, Count)
	return retVal


# Tool: 436
@mcp.tool()
async def word_Selection_Expand(this_Selection_wordObjId: str, Unit):
	"""This tool calls the Expand methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	retVal = this_Selection.Expand(Unit)
	return retVal


# Tool: 437
@mcp.tool()
async def word_Selection_InsertParagraph(this_Selection_wordObjId: str):
	"""This tool calls the InsertParagraph methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.InsertParagraph()


# Tool: 438
@mcp.tool()
async def word_Selection_InsertParagraphAfter(this_Selection_wordObjId: str):
	"""This tool calls the InsertParagraphAfter methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.InsertParagraphAfter()


# Tool: 439
@mcp.tool()
async def word_Selection_ConvertToTableOld(this_Selection_wordObjId: str, Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit):
	"""This tool calls the ConvertToTableOld methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Separator: the Separator as VT_VARIANT
		NumRows: the NumRows as VT_VARIANT
		NumColumns: the NumColumns as VT_VARIANT
		InitialColumnWidth: the InitialColumnWidth as VT_VARIANT
		Format: the Format as VT_VARIANT
		ApplyBorders: the ApplyBorders as VT_VARIANT
		ApplyShading: the ApplyShading as VT_VARIANT
		ApplyFont: the ApplyFont as VT_VARIANT
		ApplyColor: the ApplyColor as VT_VARIANT
		ApplyHeadingRows: the ApplyHeadingRows as VT_VARIANT
		ApplyLastRow: the ApplyLastRow as VT_VARIANT
		ApplyFirstColumn: the ApplyFirstColumn as VT_VARIANT
		ApplyLastColumn: the ApplyLastColumn as VT_VARIANT
		AutoFit: the AutoFit as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Separator = tryParseString(Separator)
	NumRows = tryParseString(NumRows)
	NumColumns = tryParseString(NumColumns)
	InitialColumnWidth = tryParseString(InitialColumnWidth)
	Format = tryParseString(Format)
	ApplyBorders = tryParseString(ApplyBorders)
	ApplyShading = tryParseString(ApplyShading)
	ApplyFont = tryParseString(ApplyFont)
	ApplyColor = tryParseString(ApplyColor)
	ApplyHeadingRows = tryParseString(ApplyHeadingRows)
	ApplyLastRow = tryParseString(ApplyLastRow)
	ApplyFirstColumn = tryParseString(ApplyFirstColumn)
	ApplyLastColumn = tryParseString(ApplyLastColumn)
	AutoFit = tryParseString(AutoFit)
	retVal = this_Selection.ConvertToTableOld(Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit)
	try:
		local_Uniform = retVal.Uniform
	except:
		local_Uniform = None
	try:
		local_AutoFormatType = retVal.AutoFormatType
	except:
		local_AutoFormatType = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_AllowPageBreaks = retVal.AllowPageBreaks
	except:
		local_AllowPageBreaks = None
	try:
		local_AllowAutoFit = retVal.AllowAutoFit
	except:
		local_AllowAutoFit = None
	try:
		local_PreferredWidth = retVal.PreferredWidth
	except:
		local_PreferredWidth = None
	try:
		local_PreferredWidthType = retVal.PreferredWidthType
	except:
		local_PreferredWidthType = None
	try:
		local_TopPadding = retVal.TopPadding
	except:
		local_TopPadding = None
	try:
		local_BottomPadding = retVal.BottomPadding
	except:
		local_BottomPadding = None
	try:
		local_LeftPadding = retVal.LeftPadding
	except:
		local_LeftPadding = None
	try:
		local_RightPadding = retVal.RightPadding
	except:
		local_RightPadding = None
	try:
		local_Spacing = retVal.Spacing
	except:
		local_Spacing = None
	try:
		local_TableDirection = retVal.TableDirection
	except:
		local_TableDirection = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_ApplyStyleHeadingRows = retVal.ApplyStyleHeadingRows
	except:
		local_ApplyStyleHeadingRows = None
	try:
		local_ApplyStyleLastRow = retVal.ApplyStyleLastRow
	except:
		local_ApplyStyleLastRow = None
	try:
		local_ApplyStyleFirstColumn = retVal.ApplyStyleFirstColumn
	except:
		local_ApplyStyleFirstColumn = None
	try:
		local_ApplyStyleLastColumn = retVal.ApplyStyleLastColumn
	except:
		local_ApplyStyleLastColumn = None
	try:
		local_ApplyStyleRowBands = retVal.ApplyStyleRowBands
	except:
		local_ApplyStyleRowBands = None
	try:
		local_ApplyStyleColumnBands = retVal.ApplyStyleColumnBands
	except:
		local_ApplyStyleColumnBands = None
	try:
		local_Title = retVal.Title
	except:
		local_Title = None
	try:
		local_Descr = retVal.Descr
	except:
		local_Descr = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Table", "Uniform": local_Uniform, "AutoFormatType": local_AutoFormatType, "NestingLevel": local_NestingLevel, "AllowPageBreaks": local_AllowPageBreaks, "AllowAutoFit": local_AllowAutoFit, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, "TopPadding": local_TopPadding, "BottomPadding": local_BottomPadding, "LeftPadding": local_LeftPadding, "RightPadding": local_RightPadding, "Spacing": local_Spacing, "TableDirection": local_TableDirection, "ID": local_ID, "Style": local_Style, "ApplyStyleHeadingRows": local_ApplyStyleHeadingRows, "ApplyStyleLastRow": local_ApplyStyleLastRow, "ApplyStyleFirstColumn": local_ApplyStyleFirstColumn, "ApplyStyleLastColumn": local_ApplyStyleLastColumn, "ApplyStyleRowBands": local_ApplyStyleRowBands, "ApplyStyleColumnBands": local_ApplyStyleColumnBands, "Title": local_Title, "Descr": local_Descr, }


# Tool: 440
@mcp.tool()
async def word_Selection_InsertDateTimeOld(this_Selection_wordObjId: str, DateTimeFormat, InsertAsField, InsertAsFullWidth):
	"""This tool calls the InsertDateTimeOld methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		DateTimeFormat: the DateTimeFormat as VT_VARIANT
		InsertAsField: the InsertAsField as VT_VARIANT
		InsertAsFullWidth: the InsertAsFullWidth as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	DateTimeFormat = tryParseString(DateTimeFormat)
	InsertAsField = tryParseString(InsertAsField)
	InsertAsFullWidth = tryParseString(InsertAsFullWidth)
	this_Selection.InsertDateTimeOld(DateTimeFormat, InsertAsField, InsertAsFullWidth)


# Tool: 441
@mcp.tool()
async def word_Selection_InsertSymbol(this_Selection_wordObjId: str, CharacterNumber: int, Font, Unicode, Bias):
	"""This tool calls the InsertSymbol methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		CharacterNumber: the CharacterNumber as int
		Font: the Font as VT_VARIANT
		Unicode: the Unicode as VT_VARIANT
		Bias: the Bias as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Font = tryParseString(Font)
	Unicode = tryParseString(Unicode)
	Bias = tryParseString(Bias)
	this_Selection.InsertSymbol(CharacterNumber, Font, Unicode, Bias)


# Tool: 442
@mcp.tool()
async def word_Selection_InsertCrossReference_2002(this_Selection_wordObjId: str, ReferenceType, ReferenceKind: int, ReferenceItem, InsertAsHyperlink, IncludePosition):
	"""This tool calls the InsertCrossReference_2002 methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		ReferenceType: the ReferenceType as VT_VARIANT
		ReferenceKind: the ReferenceKind as WdReferenceKind
		ReferenceItem: the ReferenceItem as VT_VARIANT
		InsertAsHyperlink: the InsertAsHyperlink as VT_VARIANT
		IncludePosition: the IncludePosition as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	ReferenceType = tryParseString(ReferenceType)
	ReferenceItem = tryParseString(ReferenceItem)
	InsertAsHyperlink = tryParseString(InsertAsHyperlink)
	IncludePosition = tryParseString(IncludePosition)
	this_Selection.InsertCrossReference_2002(ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition)


# Tool: 443
@mcp.tool()
async def word_Selection_InsertCaptionXP(this_Selection_wordObjId: str, Label, Title, TitleAutoText, Position):
	"""This tool calls the InsertCaptionXP methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Label: the Label as VT_VARIANT
		Title: the Title as VT_VARIANT
		TitleAutoText: the TitleAutoText as VT_VARIANT
		Position: the Position as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Label = tryParseString(Label)
	Title = tryParseString(Title)
	TitleAutoText = tryParseString(TitleAutoText)
	Position = tryParseString(Position)
	this_Selection.InsertCaptionXP(Label, Title, TitleAutoText, Position)


# Tool: 444
@mcp.tool()
async def word_Selection_CopyAsPicture(this_Selection_wordObjId: str):
	"""This tool calls the CopyAsPicture methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.CopyAsPicture()


# Tool: 445
@mcp.tool()
async def word_Selection_SortOld(this_Selection_wordObjId: str, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, LanguageID):
	"""This tool calls the SortOld methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		ExcludeHeader: the ExcludeHeader as VT_VARIANT
		FieldNumber: the FieldNumber as VT_VARIANT
		SortFieldType: the SortFieldType as VT_VARIANT
		SortOrder: the SortOrder as VT_VARIANT
		FieldNumber2: the FieldNumber2 as VT_VARIANT
		SortFieldType2: the SortFieldType2 as VT_VARIANT
		SortOrder2: the SortOrder2 as VT_VARIANT
		FieldNumber3: the FieldNumber3 as VT_VARIANT
		SortFieldType3: the SortFieldType3 as VT_VARIANT
		SortOrder3: the SortOrder3 as VT_VARIANT
		SortColumn: the SortColumn as VT_VARIANT
		Separator: the Separator as VT_VARIANT
		CaseSensitive: the CaseSensitive as VT_VARIANT
		LanguageID: the LanguageID as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	ExcludeHeader = tryParseString(ExcludeHeader)
	FieldNumber = tryParseString(FieldNumber)
	SortFieldType = tryParseString(SortFieldType)
	SortOrder = tryParseString(SortOrder)
	FieldNumber2 = tryParseString(FieldNumber2)
	SortFieldType2 = tryParseString(SortFieldType2)
	SortOrder2 = tryParseString(SortOrder2)
	FieldNumber3 = tryParseString(FieldNumber3)
	SortFieldType3 = tryParseString(SortFieldType3)
	SortOrder3 = tryParseString(SortOrder3)
	SortColumn = tryParseString(SortColumn)
	Separator = tryParseString(Separator)
	CaseSensitive = tryParseString(CaseSensitive)
	LanguageID = tryParseString(LanguageID)
	this_Selection.SortOld(ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, LanguageID)


# Tool: 446
@mcp.tool()
async def word_Selection_SortAscending(this_Selection_wordObjId: str):
	"""This tool calls the SortAscending methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.SortAscending()


# Tool: 447
@mcp.tool()
async def word_Selection_SortDescending(this_Selection_wordObjId: str):
	"""This tool calls the SortDescending methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.SortDescending()


# Tool: 448
@mcp.tool()
async def word_Selection_IsEqual(this_Selection_wordObjId: str, Range_wordObjId: str):
	"""This tool calls the IsEqual methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Range_wordObjId: 		To pass this object, send in the __WordObjectId of the Range object as was obtained from a previous return value
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Range = get_object(Range_wordObjId)
	retVal = this_Selection.IsEqual(Range)
	return retVal


# Tool: 449
@mcp.tool()
async def word_Selection_Calculate(this_Selection_wordObjId: str):
	"""This tool calls the Calculate methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	retVal = this_Selection.Calculate()
	return retVal


# Tool: 450
@mcp.tool()
async def word_Selection_GoTo(this_Selection_wordObjId: str, What, Which, Count, Name):
	"""This tool calls the GoTo methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		What: the What as VT_VARIANT
		Which: the Which as VT_VARIANT
		Count: the Count as VT_VARIANT
		Name: the Name as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	What = tryParseString(What)
	Which = tryParseString(Which)
	Count = tryParseString(Count)
	Name = tryParseString(Name)
	retVal = this_Selection.GoTo(What, Which, Count, Name)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 451
@mcp.tool()
async def word_Selection_GoToNext(this_Selection_wordObjId: str, What: int):
	"""This tool calls the GoToNext methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		What: the What as WdGoToItem
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	retVal = this_Selection.GoToNext(What)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 452
@mcp.tool()
async def word_Selection_GoToPrevious(this_Selection_wordObjId: str, What: int):
	"""This tool calls the GoToPrevious methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		What: the What as WdGoToItem
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	retVal = this_Selection.GoToPrevious(What)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 453
@mcp.tool()
async def word_Selection_PasteSpecial(this_Selection_wordObjId: str, IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel):
	"""This tool calls the PasteSpecial methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		IconIndex: the IconIndex as VT_VARIANT
		Link: the Link as VT_VARIANT
		Placement: the Placement as VT_VARIANT
		DisplayAsIcon: the DisplayAsIcon as VT_VARIANT
		DataType: the DataType as VT_VARIANT
		IconFileName: the IconFileName as VT_VARIANT
		IconLabel: the IconLabel as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	IconIndex = tryParseString(IconIndex)
	Link = tryParseString(Link)
	Placement = tryParseString(Placement)
	DisplayAsIcon = tryParseString(DisplayAsIcon)
	DataType = tryParseString(DataType)
	IconFileName = tryParseString(IconFileName)
	IconLabel = tryParseString(IconLabel)
	this_Selection.PasteSpecial(IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel)


# Tool: 454
@mcp.tool()
async def word_Selection_PreviousField(this_Selection_wordObjId: str):
	"""This tool calls the PreviousField methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	retVal = this_Selection.PreviousField()
	try:
		local_Type = retVal.Type
	except:
		local_Type = None
	try:
		local_Locked = retVal.Locked
	except:
		local_Locked = None
	try:
		local_Kind = retVal.Kind
	except:
		local_Kind = None
	try:
		local_Data = retVal.Data
	except:
		local_Data = None
	try:
		local_Index = retVal.Index
	except:
		local_Index = None
	try:
		local_ShowCodes = retVal.ShowCodes
	except:
		local_ShowCodes = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Field", "Type": local_Type, "Locked": local_Locked, "Kind": local_Kind, "Data": local_Data, "Index": local_Index, "ShowCodes": local_ShowCodes, }


# Tool: 455
@mcp.tool()
async def word_Selection_NextField(this_Selection_wordObjId: str):
	"""This tool calls the NextField methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	retVal = this_Selection.NextField()
	try:
		local_Type = retVal.Type
	except:
		local_Type = None
	try:
		local_Locked = retVal.Locked
	except:
		local_Locked = None
	try:
		local_Kind = retVal.Kind
	except:
		local_Kind = None
	try:
		local_Data = retVal.Data
	except:
		local_Data = None
	try:
		local_Index = retVal.Index
	except:
		local_Index = None
	try:
		local_ShowCodes = retVal.ShowCodes
	except:
		local_ShowCodes = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Field", "Type": local_Type, "Locked": local_Locked, "Kind": local_Kind, "Data": local_Data, "Index": local_Index, "ShowCodes": local_ShowCodes, }


# Tool: 456
@mcp.tool()
async def word_Selection_InsertParagraphBefore(this_Selection_wordObjId: str):
	"""This tool calls the InsertParagraphBefore methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.InsertParagraphBefore()


# Tool: 457
@mcp.tool()
async def word_Selection_InsertCells(this_Selection_wordObjId: str, ShiftCells):
	"""This tool calls the InsertCells methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		ShiftCells: the ShiftCells as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	ShiftCells = tryParseString(ShiftCells)
	this_Selection.InsertCells(ShiftCells)


# Tool: 458
@mcp.tool()
async def word_Selection_Extend(this_Selection_wordObjId: str, Character):
	"""This tool calls the Extend methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Character: the Character as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Character = tryParseString(Character)
	this_Selection.Extend(Character)


# Tool: 459
@mcp.tool()
async def word_Selection_Shrink(this_Selection_wordObjId: str):
	"""This tool calls the Shrink methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.Shrink()


# Tool: 460
@mcp.tool()
async def word_Selection_MoveLeft(this_Selection_wordObjId: str, Unit, Count, Extend):
	"""This tool calls the MoveLeft methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
		Extend: the Extend as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	Extend = tryParseString(Extend)
	retVal = this_Selection.MoveLeft(Unit, Count, Extend)
	return retVal


# Tool: 461
@mcp.tool()
async def word_Selection_MoveRight(this_Selection_wordObjId: str, Unit, Count, Extend):
	"""This tool calls the MoveRight methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
		Extend: the Extend as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	Extend = tryParseString(Extend)
	retVal = this_Selection.MoveRight(Unit, Count, Extend)
	return retVal


# Tool: 462
@mcp.tool()
async def word_Selection_MoveUp(this_Selection_wordObjId: str, Unit, Count, Extend):
	"""This tool calls the MoveUp methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
		Extend: the Extend as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	Extend = tryParseString(Extend)
	retVal = this_Selection.MoveUp(Unit, Count, Extend)
	return retVal


# Tool: 463
@mcp.tool()
async def word_Selection_MoveDown(this_Selection_wordObjId: str, Unit, Count, Extend):
	"""This tool calls the MoveDown methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Count: the Count as VT_VARIANT
		Extend: the Extend as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Count = tryParseString(Count)
	Extend = tryParseString(Extend)
	retVal = this_Selection.MoveDown(Unit, Count, Extend)
	return retVal


# Tool: 464
@mcp.tool()
async def word_Selection_HomeKey(this_Selection_wordObjId: str, Unit, Extend):
	"""This tool calls the HomeKey methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Extend: the Extend as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Extend = tryParseString(Extend)
	retVal = this_Selection.HomeKey(Unit, Extend)
	return retVal


# Tool: 465
@mcp.tool()
async def word_Selection_EndKey(this_Selection_wordObjId: str, Unit, Extend):
	"""This tool calls the EndKey methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Unit: the Unit as VT_VARIANT
		Extend: the Extend as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Unit = tryParseString(Unit)
	Extend = tryParseString(Extend)
	retVal = this_Selection.EndKey(Unit, Extend)
	return retVal


# Tool: 466
@mcp.tool()
async def word_Selection_EscapeKey(this_Selection_wordObjId: str):
	"""This tool calls the EscapeKey methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.EscapeKey()


# Tool: 467
@mcp.tool()
async def word_Selection_TypeText(this_Selection_wordObjId: str, Text: str):
	"""This tool calls the TypeText methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Text: the Text as str
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.TypeText(Text)


# Tool: 468
@mcp.tool()
async def word_Selection_CopyFormat(this_Selection_wordObjId: str):
	"""This tool calls the CopyFormat methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.CopyFormat()


# Tool: 469
@mcp.tool()
async def word_Selection_PasteFormat(this_Selection_wordObjId: str):
	"""This tool calls the PasteFormat methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.PasteFormat()


# Tool: 470
@mcp.tool()
async def word_Selection_TypeParagraph(this_Selection_wordObjId: str):
	"""This tool calls the TypeParagraph methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.TypeParagraph()


# Tool: 471
@mcp.tool()
async def word_Selection_TypeBackspace(this_Selection_wordObjId: str):
	"""This tool calls the TypeBackspace methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.TypeBackspace()


# Tool: 472
@mcp.tool()
async def word_Selection_NextSubdocument(this_Selection_wordObjId: str):
	"""This tool calls the NextSubdocument methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.NextSubdocument()


# Tool: 473
@mcp.tool()
async def word_Selection_PreviousSubdocument(this_Selection_wordObjId: str):
	"""This tool calls the PreviousSubdocument methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.PreviousSubdocument()


# Tool: 474
@mcp.tool()
async def word_Selection_SelectColumn(this_Selection_wordObjId: str):
	"""This tool calls the SelectColumn methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.SelectColumn()


# Tool: 475
@mcp.tool()
async def word_Selection_SelectCurrentFont(this_Selection_wordObjId: str):
	"""This tool calls the SelectCurrentFont methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.SelectCurrentFont()


# Tool: 476
@mcp.tool()
async def word_Selection_SelectCurrentAlignment(this_Selection_wordObjId: str):
	"""This tool calls the SelectCurrentAlignment methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.SelectCurrentAlignment()


# Tool: 477
@mcp.tool()
async def word_Selection_SelectCurrentSpacing(this_Selection_wordObjId: str):
	"""This tool calls the SelectCurrentSpacing methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.SelectCurrentSpacing()


# Tool: 478
@mcp.tool()
async def word_Selection_SelectCurrentIndent(this_Selection_wordObjId: str):
	"""This tool calls the SelectCurrentIndent methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.SelectCurrentIndent()


# Tool: 479
@mcp.tool()
async def word_Selection_SelectCurrentTabs(this_Selection_wordObjId: str):
	"""This tool calls the SelectCurrentTabs methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.SelectCurrentTabs()


# Tool: 480
@mcp.tool()
async def word_Selection_SelectCurrentColor(this_Selection_wordObjId: str):
	"""This tool calls the SelectCurrentColor methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.SelectCurrentColor()


# Tool: 481
@mcp.tool()
async def word_Selection_CreateTextbox(this_Selection_wordObjId: str):
	"""This tool calls the CreateTextbox methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.CreateTextbox()


# Tool: 482
@mcp.tool()
async def word_Selection_WholeStory(this_Selection_wordObjId: str):
	"""This tool calls the WholeStory methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.WholeStory()


# Tool: 483
@mcp.tool()
async def word_Selection_SelectRow(this_Selection_wordObjId: str):
	"""This tool calls the SelectRow methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.SelectRow()


# Tool: 484
@mcp.tool()
async def word_Selection_SplitTable(this_Selection_wordObjId: str):
	"""This tool calls the SplitTable methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.SplitTable()


# Tool: 485
@mcp.tool()
async def word_Selection_InsertRows(this_Selection_wordObjId: str, NumRows):
	"""This tool calls the InsertRows methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		NumRows: the NumRows as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	NumRows = tryParseString(NumRows)
	this_Selection.InsertRows(NumRows)


# Tool: 486
@mcp.tool()
async def word_Selection_InsertColumns(this_Selection_wordObjId: str):
	"""This tool calls the InsertColumns methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.InsertColumns()


# Tool: 487
@mcp.tool()
async def word_Selection_InsertFormula(this_Selection_wordObjId: str, Formula, NumberFormat):
	"""This tool calls the InsertFormula methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Formula: the Formula as VT_VARIANT
		NumberFormat: the NumberFormat as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Formula = tryParseString(Formula)
	NumberFormat = tryParseString(NumberFormat)
	this_Selection.InsertFormula(Formula, NumberFormat)


# Tool: 488
@mcp.tool()
async def word_Selection_NextRevision(this_Selection_wordObjId: str, Wrap):
	"""This tool calls the NextRevision methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Wrap: the Wrap as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Wrap = tryParseString(Wrap)
	retVal = this_Selection.NextRevision(Wrap)
	try:
		local_Author = retVal.Author
	except:
		local_Author = None
	try:
		local_Date = retVal.Date
	except:
		local_Date = None
	try:
		local_Type = retVal.Type
	except:
		local_Type = None
	try:
		local_Index = retVal.Index
	except:
		local_Index = None
	try:
		local_FormatDescription = retVal.FormatDescription
	except:
		local_FormatDescription = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Revision", "Author": local_Author, "Date": local_Date, "Type": local_Type, "Index": local_Index, "FormatDescription": local_FormatDescription, }


# Tool: 489
@mcp.tool()
async def word_Selection_PreviousRevision(this_Selection_wordObjId: str, Wrap):
	"""This tool calls the PreviousRevision methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Wrap: the Wrap as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Wrap = tryParseString(Wrap)
	retVal = this_Selection.PreviousRevision(Wrap)
	try:
		local_Author = retVal.Author
	except:
		local_Author = None
	try:
		local_Date = retVal.Date
	except:
		local_Date = None
	try:
		local_Type = retVal.Type
	except:
		local_Type = None
	try:
		local_Index = retVal.Index
	except:
		local_Index = None
	try:
		local_FormatDescription = retVal.FormatDescription
	except:
		local_FormatDescription = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Revision", "Author": local_Author, "Date": local_Date, "Type": local_Type, "Index": local_Index, "FormatDescription": local_FormatDescription, }


# Tool: 490
@mcp.tool()
async def word_Selection_PasteAsNestedTable(this_Selection_wordObjId: str):
	"""This tool calls the PasteAsNestedTable methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.PasteAsNestedTable()


# Tool: 491
@mcp.tool()
async def word_Selection_CreateAutoTextEntry(this_Selection_wordObjId: str, Name: str, StyleName: str):
	"""This tool calls the CreateAutoTextEntry methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Name: the Name as str
		StyleName: the StyleName as str
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	retVal = this_Selection.CreateAutoTextEntry(Name, StyleName)
	try:
		local_Index = retVal.Index
	except:
		local_Index = None
	try:
		local_Name = retVal.Name
	except:
		local_Name = None
	try:
		local_StyleName = retVal.StyleName
	except:
		local_StyleName = None
	try:
		local_Value = retVal.Value
	except:
		local_Value = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "AutoTextEntry", "Index": local_Index, "Name": local_Name, "StyleName": local_StyleName, "Value": local_Value, }


# Tool: 492
@mcp.tool()
async def word_Selection_DetectLanguage(this_Selection_wordObjId: str):
	"""This tool calls the DetectLanguage methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.DetectLanguage()


# Tool: 493
@mcp.tool()
async def word_Selection_SelectCell(this_Selection_wordObjId: str):
	"""This tool calls the SelectCell methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.SelectCell()


# Tool: 494
@mcp.tool()
async def word_Selection_InsertRowsBelow(this_Selection_wordObjId: str, NumRows):
	"""This tool calls the InsertRowsBelow methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		NumRows: the NumRows as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	NumRows = tryParseString(NumRows)
	this_Selection.InsertRowsBelow(NumRows)


# Tool: 495
@mcp.tool()
async def word_Selection_InsertColumnsRight(this_Selection_wordObjId: str):
	"""This tool calls the InsertColumnsRight methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.InsertColumnsRight()


# Tool: 496
@mcp.tool()
async def word_Selection_InsertRowsAbove(this_Selection_wordObjId: str, NumRows):
	"""This tool calls the InsertRowsAbove methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		NumRows: the NumRows as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	NumRows = tryParseString(NumRows)
	this_Selection.InsertRowsAbove(NumRows)


# Tool: 497
@mcp.tool()
async def word_Selection_RtlRun(this_Selection_wordObjId: str):
	"""This tool calls the RtlRun methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.RtlRun()


# Tool: 498
@mcp.tool()
async def word_Selection_LtrRun(this_Selection_wordObjId: str):
	"""This tool calls the LtrRun methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.LtrRun()


# Tool: 499
@mcp.tool()
async def word_Selection_BoldRun(this_Selection_wordObjId: str):
	"""This tool calls the BoldRun methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.BoldRun()


# Tool: 500
@mcp.tool()
async def word_Selection_ItalicRun(this_Selection_wordObjId: str):
	"""This tool calls the ItalicRun methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.ItalicRun()


# Tool: 501
@mcp.tool()
async def word_Selection_RtlPara(this_Selection_wordObjId: str):
	"""This tool calls the RtlPara methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.RtlPara()


# Tool: 502
@mcp.tool()
async def word_Selection_LtrPara(this_Selection_wordObjId: str):
	"""This tool calls the LtrPara methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.LtrPara()


# Tool: 503
@mcp.tool()
async def word_Selection_InsertDateTime(this_Selection_wordObjId: str, DateTimeFormat, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType):
	"""This tool calls the InsertDateTime methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		DateTimeFormat: the DateTimeFormat as VT_VARIANT
		InsertAsField: the InsertAsField as VT_VARIANT
		InsertAsFullWidth: the InsertAsFullWidth as VT_VARIANT
		DateLanguage: the DateLanguage as VT_VARIANT
		CalendarType: the CalendarType as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	DateTimeFormat = tryParseString(DateTimeFormat)
	InsertAsField = tryParseString(InsertAsField)
	InsertAsFullWidth = tryParseString(InsertAsFullWidth)
	DateLanguage = tryParseString(DateLanguage)
	CalendarType = tryParseString(CalendarType)
	this_Selection.InsertDateTime(DateTimeFormat, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType)


# Tool: 504
@mcp.tool()
async def word_Selection_Sort2000(this_Selection_wordObjId: str, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID):
	"""This tool calls the Sort2000 methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		ExcludeHeader: the ExcludeHeader as VT_VARIANT
		FieldNumber: the FieldNumber as VT_VARIANT
		SortFieldType: the SortFieldType as VT_VARIANT
		SortOrder: the SortOrder as VT_VARIANT
		FieldNumber2: the FieldNumber2 as VT_VARIANT
		SortFieldType2: the SortFieldType2 as VT_VARIANT
		SortOrder2: the SortOrder2 as VT_VARIANT
		FieldNumber3: the FieldNumber3 as VT_VARIANT
		SortFieldType3: the SortFieldType3 as VT_VARIANT
		SortOrder3: the SortOrder3 as VT_VARIANT
		SortColumn: the SortColumn as VT_VARIANT
		Separator: the Separator as VT_VARIANT
		CaseSensitive: the CaseSensitive as VT_VARIANT
		BidiSort: the BidiSort as VT_VARIANT
		IgnoreThe: the IgnoreThe as VT_VARIANT
		IgnoreKashida: the IgnoreKashida as VT_VARIANT
		IgnoreDiacritics: the IgnoreDiacritics as VT_VARIANT
		IgnoreHe: the IgnoreHe as VT_VARIANT
		LanguageID: the LanguageID as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	ExcludeHeader = tryParseString(ExcludeHeader)
	FieldNumber = tryParseString(FieldNumber)
	SortFieldType = tryParseString(SortFieldType)
	SortOrder = tryParseString(SortOrder)
	FieldNumber2 = tryParseString(FieldNumber2)
	SortFieldType2 = tryParseString(SortFieldType2)
	SortOrder2 = tryParseString(SortOrder2)
	FieldNumber3 = tryParseString(FieldNumber3)
	SortFieldType3 = tryParseString(SortFieldType3)
	SortOrder3 = tryParseString(SortOrder3)
	SortColumn = tryParseString(SortColumn)
	Separator = tryParseString(Separator)
	CaseSensitive = tryParseString(CaseSensitive)
	BidiSort = tryParseString(BidiSort)
	IgnoreThe = tryParseString(IgnoreThe)
	IgnoreKashida = tryParseString(IgnoreKashida)
	IgnoreDiacritics = tryParseString(IgnoreDiacritics)
	IgnoreHe = tryParseString(IgnoreHe)
	LanguageID = tryParseString(LanguageID)
	this_Selection.Sort2000(ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID)


# Tool: 505
@mcp.tool()
async def word_Selection_ConvertToTable(this_Selection_wordObjId: str, Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior):
	"""This tool calls the ConvertToTable methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Separator: the Separator as VT_VARIANT
		NumRows: the NumRows as VT_VARIANT
		NumColumns: the NumColumns as VT_VARIANT
		InitialColumnWidth: the InitialColumnWidth as VT_VARIANT
		Format: the Format as VT_VARIANT
		ApplyBorders: the ApplyBorders as VT_VARIANT
		ApplyShading: the ApplyShading as VT_VARIANT
		ApplyFont: the ApplyFont as VT_VARIANT
		ApplyColor: the ApplyColor as VT_VARIANT
		ApplyHeadingRows: the ApplyHeadingRows as VT_VARIANT
		ApplyLastRow: the ApplyLastRow as VT_VARIANT
		ApplyFirstColumn: the ApplyFirstColumn as VT_VARIANT
		ApplyLastColumn: the ApplyLastColumn as VT_VARIANT
		AutoFit: the AutoFit as VT_VARIANT
		AutoFitBehavior: the AutoFitBehavior as VT_VARIANT
		DefaultTableBehavior: the DefaultTableBehavior as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Separator = tryParseString(Separator)
	NumRows = tryParseString(NumRows)
	NumColumns = tryParseString(NumColumns)
	InitialColumnWidth = tryParseString(InitialColumnWidth)
	Format = tryParseString(Format)
	ApplyBorders = tryParseString(ApplyBorders)
	ApplyShading = tryParseString(ApplyShading)
	ApplyFont = tryParseString(ApplyFont)
	ApplyColor = tryParseString(ApplyColor)
	ApplyHeadingRows = tryParseString(ApplyHeadingRows)
	ApplyLastRow = tryParseString(ApplyLastRow)
	ApplyFirstColumn = tryParseString(ApplyFirstColumn)
	ApplyLastColumn = tryParseString(ApplyLastColumn)
	AutoFit = tryParseString(AutoFit)
	AutoFitBehavior = tryParseString(AutoFitBehavior)
	DefaultTableBehavior = tryParseString(DefaultTableBehavior)
	retVal = this_Selection.ConvertToTable(Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior)
	try:
		local_Uniform = retVal.Uniform
	except:
		local_Uniform = None
	try:
		local_AutoFormatType = retVal.AutoFormatType
	except:
		local_AutoFormatType = None
	try:
		local_NestingLevel = retVal.NestingLevel
	except:
		local_NestingLevel = None
	try:
		local_AllowPageBreaks = retVal.AllowPageBreaks
	except:
		local_AllowPageBreaks = None
	try:
		local_AllowAutoFit = retVal.AllowAutoFit
	except:
		local_AllowAutoFit = None
	try:
		local_PreferredWidth = retVal.PreferredWidth
	except:
		local_PreferredWidth = None
	try:
		local_PreferredWidthType = retVal.PreferredWidthType
	except:
		local_PreferredWidthType = None
	try:
		local_TopPadding = retVal.TopPadding
	except:
		local_TopPadding = None
	try:
		local_BottomPadding = retVal.BottomPadding
	except:
		local_BottomPadding = None
	try:
		local_LeftPadding = retVal.LeftPadding
	except:
		local_LeftPadding = None
	try:
		local_RightPadding = retVal.RightPadding
	except:
		local_RightPadding = None
	try:
		local_Spacing = retVal.Spacing
	except:
		local_Spacing = None
	try:
		local_TableDirection = retVal.TableDirection
	except:
		local_TableDirection = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_ApplyStyleHeadingRows = retVal.ApplyStyleHeadingRows
	except:
		local_ApplyStyleHeadingRows = None
	try:
		local_ApplyStyleLastRow = retVal.ApplyStyleLastRow
	except:
		local_ApplyStyleLastRow = None
	try:
		local_ApplyStyleFirstColumn = retVal.ApplyStyleFirstColumn
	except:
		local_ApplyStyleFirstColumn = None
	try:
		local_ApplyStyleLastColumn = retVal.ApplyStyleLastColumn
	except:
		local_ApplyStyleLastColumn = None
	try:
		local_ApplyStyleRowBands = retVal.ApplyStyleRowBands
	except:
		local_ApplyStyleRowBands = None
	try:
		local_ApplyStyleColumnBands = retVal.ApplyStyleColumnBands
	except:
		local_ApplyStyleColumnBands = None
	try:
		local_Title = retVal.Title
	except:
		local_Title = None
	try:
		local_Descr = retVal.Descr
	except:
		local_Descr = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Table", "Uniform": local_Uniform, "AutoFormatType": local_AutoFormatType, "NestingLevel": local_NestingLevel, "AllowPageBreaks": local_AllowPageBreaks, "AllowAutoFit": local_AllowAutoFit, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, "TopPadding": local_TopPadding, "BottomPadding": local_BottomPadding, "LeftPadding": local_LeftPadding, "RightPadding": local_RightPadding, "Spacing": local_Spacing, "TableDirection": local_TableDirection, "ID": local_ID, "Style": local_Style, "ApplyStyleHeadingRows": local_ApplyStyleHeadingRows, "ApplyStyleLastRow": local_ApplyStyleLastRow, "ApplyStyleFirstColumn": local_ApplyStyleFirstColumn, "ApplyStyleLastColumn": local_ApplyStyleLastColumn, "ApplyStyleRowBands": local_ApplyStyleRowBands, "ApplyStyleColumnBands": local_ApplyStyleColumnBands, "Title": local_Title, "Descr": local_Descr, }


# Tool: 506
@mcp.tool()
async def word_Selection_ClearFormatting(this_Selection_wordObjId: str):
	"""This tool calls the ClearFormatting methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.ClearFormatting()


# Tool: 507
@mcp.tool()
async def word_Selection_PasteAppendTable(this_Selection_wordObjId: str):
	"""This tool calls the PasteAppendTable methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.PasteAppendTable()


# Tool: 508
@mcp.tool()
async def word_Selection_ToggleCharacterCode(this_Selection_wordObjId: str):
	"""This tool calls the ToggleCharacterCode methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.ToggleCharacterCode()


# Tool: 509
@mcp.tool()
async def word_Selection_PasteAndFormat(this_Selection_wordObjId: str, Type: int):
	"""This tool calls the PasteAndFormat methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Type: the Type as WdRecoveryType
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.PasteAndFormat(Type)


# Tool: 510
@mcp.tool()
async def word_Selection_PasteExcelTable(this_Selection_wordObjId: str, LinkedToExcel: bool, WordFormatting: bool, RTF: bool):
	"""This tool calls the PasteExcelTable methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		LinkedToExcel: the LinkedToExcel as bool
		WordFormatting: the WordFormatting as bool
		RTF: the RTF as bool
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.PasteExcelTable(LinkedToExcel, WordFormatting, RTF)


# Tool: 511
@mcp.tool()
async def word_Selection_ShrinkDiscontiguousSelection(this_Selection_wordObjId: str):
	"""This tool calls the ShrinkDiscontiguousSelection methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.ShrinkDiscontiguousSelection()


# Tool: 512
@mcp.tool()
async def word_Selection_InsertStyleSeparator(this_Selection_wordObjId: str):
	"""This tool calls the InsertStyleSeparator methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.InsertStyleSeparator()


# Tool: 513
@mcp.tool()
async def word_Selection_Sort(this_Selection_wordObjId: str, ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID, SubFieldNumber, SubFieldNumber2, SubFieldNumber3):
	"""This tool calls the Sort methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		ExcludeHeader: the ExcludeHeader as VT_VARIANT
		FieldNumber: the FieldNumber as VT_VARIANT
		SortFieldType: the SortFieldType as VT_VARIANT
		SortOrder: the SortOrder as VT_VARIANT
		FieldNumber2: the FieldNumber2 as VT_VARIANT
		SortFieldType2: the SortFieldType2 as VT_VARIANT
		SortOrder2: the SortOrder2 as VT_VARIANT
		FieldNumber3: the FieldNumber3 as VT_VARIANT
		SortFieldType3: the SortFieldType3 as VT_VARIANT
		SortOrder3: the SortOrder3 as VT_VARIANT
		SortColumn: the SortColumn as VT_VARIANT
		Separator: the Separator as VT_VARIANT
		CaseSensitive: the CaseSensitive as VT_VARIANT
		BidiSort: the BidiSort as VT_VARIANT
		IgnoreThe: the IgnoreThe as VT_VARIANT
		IgnoreKashida: the IgnoreKashida as VT_VARIANT
		IgnoreDiacritics: the IgnoreDiacritics as VT_VARIANT
		IgnoreHe: the IgnoreHe as VT_VARIANT
		LanguageID: the LanguageID as VT_VARIANT
		SubFieldNumber: the SubFieldNumber as VT_VARIANT
		SubFieldNumber2: the SubFieldNumber2 as VT_VARIANT
		SubFieldNumber3: the SubFieldNumber3 as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	ExcludeHeader = tryParseString(ExcludeHeader)
	FieldNumber = tryParseString(FieldNumber)
	SortFieldType = tryParseString(SortFieldType)
	SortOrder = tryParseString(SortOrder)
	FieldNumber2 = tryParseString(FieldNumber2)
	SortFieldType2 = tryParseString(SortFieldType2)
	SortOrder2 = tryParseString(SortOrder2)
	FieldNumber3 = tryParseString(FieldNumber3)
	SortFieldType3 = tryParseString(SortFieldType3)
	SortOrder3 = tryParseString(SortOrder3)
	SortColumn = tryParseString(SortColumn)
	Separator = tryParseString(Separator)
	CaseSensitive = tryParseString(CaseSensitive)
	BidiSort = tryParseString(BidiSort)
	IgnoreThe = tryParseString(IgnoreThe)
	IgnoreKashida = tryParseString(IgnoreKashida)
	IgnoreDiacritics = tryParseString(IgnoreDiacritics)
	IgnoreHe = tryParseString(IgnoreHe)
	LanguageID = tryParseString(LanguageID)
	SubFieldNumber = tryParseString(SubFieldNumber)
	SubFieldNumber2 = tryParseString(SubFieldNumber2)
	SubFieldNumber3 = tryParseString(SubFieldNumber3)
	this_Selection.Sort(ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID, SubFieldNumber, SubFieldNumber2, SubFieldNumber3)


# Tool: 514
@mcp.tool()
async def word_Selection_GoToEditableRange(this_Selection_wordObjId: str, EditorID):
	"""This tool calls the GoToEditableRange methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		EditorID: the EditorID as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	EditorID = tryParseString(EditorID)
	retVal = this_Selection.GoToEditableRange(EditorID)
	try:
		local_Text = retVal.Text
	except:
		local_Text = None
	try:
		local_Start = retVal.Start
	except:
		local_Start = None
	try:
		local_End = retVal.End
	except:
		local_End = None
	try:
		local_StoryType = retVal.StoryType
	except:
		local_StoryType = None
	try:
		local_Bold = retVal.Bold
	except:
		local_Bold = None
	try:
		local_Italic = retVal.Italic
	except:
		local_Italic = None
	try:
		local_Underline = retVal.Underline
	except:
		local_Underline = None
	try:
		local_EmphasisMark = retVal.EmphasisMark
	except:
		local_EmphasisMark = None
	try:
		local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
	except:
		local_DisableCharacterSpaceGrid = None
	try:
		local_Style = retVal.Style
	except:
		local_Style = None
	try:
		local_StoryLength = retVal.StoryLength
	except:
		local_StoryLength = None
	try:
		local_LanguageID = retVal.LanguageID
	except:
		local_LanguageID = None
	try:
		local_GrammarChecked = retVal.GrammarChecked
	except:
		local_GrammarChecked = None
	try:
		local_SpellingChecked = retVal.SpellingChecked
	except:
		local_SpellingChecked = None
	try:
		local_HighlightColorIndex = retVal.HighlightColorIndex
	except:
		local_HighlightColorIndex = None
	try:
		local_CanEdit = retVal.CanEdit
	except:
		local_CanEdit = None
	try:
		local_CanPaste = retVal.CanPaste
	except:
		local_CanPaste = None
	try:
		local_IsEndOfRowMark = retVal.IsEndOfRowMark
	except:
		local_IsEndOfRowMark = None
	try:
		local_BookmarkID = retVal.BookmarkID
	except:
		local_BookmarkID = None
	try:
		local_PreviousBookmarkID = retVal.PreviousBookmarkID
	except:
		local_PreviousBookmarkID = None
	try:
		local_Case = retVal.Case
	except:
		local_Case = None
	try:
		local_Information = retVal.Information
	except:
		local_Information = None
	try:
		local_Orientation = retVal.Orientation
	except:
		local_Orientation = None
	try:
		local_LanguageIDFarEast = retVal.LanguageIDFarEast
	except:
		local_LanguageIDFarEast = None
	try:
		local_LanguageIDOther = retVal.LanguageIDOther
	except:
		local_LanguageIDOther = None
	try:
		local_LanguageDetected = retVal.LanguageDetected
	except:
		local_LanguageDetected = None
	try:
		local_FitTextWidth = retVal.FitTextWidth
	except:
		local_FitTextWidth = None
	try:
		local_HorizontalInVertical = retVal.HorizontalInVertical
	except:
		local_HorizontalInVertical = None
	try:
		local_TwoLinesInOne = retVal.TwoLinesInOne
	except:
		local_TwoLinesInOne = None
	try:
		local_CombineCharacters = retVal.CombineCharacters
	except:
		local_CombineCharacters = None
	try:
		local_NoProofing = retVal.NoProofing
	except:
		local_NoProofing = None
	try:
		local_CharacterWidth = retVal.CharacterWidth
	except:
		local_CharacterWidth = None
	try:
		local_Kana = retVal.Kana
	except:
		local_Kana = None
	try:
		local_BoldBi = retVal.BoldBi
	except:
		local_BoldBi = None
	try:
		local_ItalicBi = retVal.ItalicBi
	except:
		local_ItalicBi = None
	try:
		local_ID = retVal.ID
	except:
		local_ID = None
	try:
		local_ShowAll = retVal.ShowAll
	except:
		local_ShowAll = None
	try:
		local_CharacterStyle = retVal.CharacterStyle
	except:
		local_CharacterStyle = None
	try:
		local_ParagraphStyle = retVal.ParagraphStyle
	except:
		local_ParagraphStyle = None
	try:
		local_ListStyle = retVal.ListStyle
	except:
		local_ListStyle = None
	try:
		local_TableStyle = retVal.TableStyle
	except:
		local_TableStyle = None
	try:
		local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
	except:
		local_TextVisibleOnScreen = None
	return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }


# Tool: 515
@mcp.tool()
async def word_Selection_InsertXML(this_Selection_wordObjId: str, XML: str, Transform):
	"""This tool calls the InsertXML methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		XML: the XML as str
		Transform: the Transform as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Transform = tryParseString(Transform)
	this_Selection.InsertXML(XML, Transform)


# Tool: 516
@mcp.tool()
async def word_Selection_InsertCaption(this_Selection_wordObjId: str, Label, Title, TitleAutoText, Position, ExcludeLabel):
	"""This tool calls the InsertCaption methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		Label: the Label as VT_VARIANT
		Title: the Title as VT_VARIANT
		TitleAutoText: the TitleAutoText as VT_VARIANT
		Position: the Position as VT_VARIANT
		ExcludeLabel: the ExcludeLabel as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	Label = tryParseString(Label)
	Title = tryParseString(Title)
	TitleAutoText = tryParseString(TitleAutoText)
	Position = tryParseString(Position)
	ExcludeLabel = tryParseString(ExcludeLabel)
	this_Selection.InsertCaption(Label, Title, TitleAutoText, Position, ExcludeLabel)


# Tool: 517
@mcp.tool()
async def word_Selection_InsertCrossReference(this_Selection_wordObjId: str, ReferenceType, ReferenceKind: int, ReferenceItem, InsertAsHyperlink, IncludePosition, SeparateNumbers, SeparatorString):
	"""This tool calls the InsertCrossReference methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		ReferenceType: the ReferenceType as VT_VARIANT
		ReferenceKind: the ReferenceKind as WdReferenceKind
		ReferenceItem: the ReferenceItem as VT_VARIANT
		InsertAsHyperlink: the InsertAsHyperlink as VT_VARIANT
		IncludePosition: the IncludePosition as VT_VARIANT
		SeparateNumbers: the SeparateNumbers as VT_VARIANT
		SeparatorString: the SeparatorString as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	ReferenceType = tryParseString(ReferenceType)
	ReferenceItem = tryParseString(ReferenceItem)
	InsertAsHyperlink = tryParseString(InsertAsHyperlink)
	IncludePosition = tryParseString(IncludePosition)
	SeparateNumbers = tryParseString(SeparateNumbers)
	SeparatorString = tryParseString(SeparatorString)
	this_Selection.InsertCrossReference(ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition, SeparateNumbers, SeparatorString)


# Tool: 518
@mcp.tool()
async def word_Selection_ClearParagraphStyle(this_Selection_wordObjId: str):
	"""This tool calls the ClearParagraphStyle methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.ClearParagraphStyle()


# Tool: 519
@mcp.tool()
async def word_Selection_ClearCharacterAllFormatting(this_Selection_wordObjId: str):
	"""This tool calls the ClearCharacterAllFormatting methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.ClearCharacterAllFormatting()


# Tool: 520
@mcp.tool()
async def word_Selection_ClearCharacterStyle(this_Selection_wordObjId: str):
	"""This tool calls the ClearCharacterStyle methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.ClearCharacterStyle()


# Tool: 521
@mcp.tool()
async def word_Selection_ClearCharacterDirectFormatting(this_Selection_wordObjId: str):
	"""This tool calls the ClearCharacterDirectFormatting methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.ClearCharacterDirectFormatting()


# Tool: 522
@mcp.tool()
async def word_Selection_ExportAsFixedFormat(this_Selection_wordObjId: str, OutputFileName: str, ExportFormat: int, OpenAfterExport: bool, OptimizeFor: int, ExportCurrentPage: bool, Item: int, IncludeDocProps: bool, KeepIRM: bool, CreateBookmarks: int, DocStructureTags: bool, BitmapMissingFonts: bool, UseISO19005_1: bool, FixedFormatExtClassPtr):
	"""This tool calls the ExportAsFixedFormat methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		OutputFileName: the OutputFileName as str
		ExportFormat: the ExportFormat as WdExportFormat
		OpenAfterExport: the OpenAfterExport as bool
		OptimizeFor: the OptimizeFor as WdExportOptimizeFor
		ExportCurrentPage: the ExportCurrentPage as bool
		Item: the Item as WdExportItem
		IncludeDocProps: the IncludeDocProps as bool
		KeepIRM: the KeepIRM as bool
		CreateBookmarks: the CreateBookmarks as WdExportCreateBookmarks
		DocStructureTags: the DocStructureTags as bool
		BitmapMissingFonts: the BitmapMissingFonts as bool
		UseISO19005_1: the UseISO19005_1 as bool
		FixedFormatExtClassPtr: the FixedFormatExtClassPtr as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	FixedFormatExtClassPtr = tryParseString(FixedFormatExtClassPtr)
	this_Selection.ExportAsFixedFormat(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr)


# Tool: 523
@mcp.tool()
async def word_Selection_ReadingModeGrowFont(this_Selection_wordObjId: str):
	"""This tool calls the ReadingModeGrowFont methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.ReadingModeGrowFont()


# Tool: 524
@mcp.tool()
async def word_Selection_ReadingModeShrinkFont(this_Selection_wordObjId: str):
	"""This tool calls the ReadingModeShrinkFont methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.ReadingModeShrinkFont()


# Tool: 525
@mcp.tool()
async def word_Selection_ClearParagraphAllFormatting(this_Selection_wordObjId: str):
	"""This tool calls the ClearParagraphAllFormatting methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.ClearParagraphAllFormatting()


# Tool: 526
@mcp.tool()
async def word_Selection_ClearParagraphDirectFormatting(this_Selection_wordObjId: str):
	"""This tool calls the ClearParagraphDirectFormatting methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.ClearParagraphDirectFormatting()


# Tool: 527
@mcp.tool()
async def word_Selection_InsertNewPage(this_Selection_wordObjId: str):
	"""This tool calls the InsertNewPage methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
"""
	this_Selection = get_object(this_Selection_wordObjId)
	this_Selection.InsertNewPage()


# Tool: 528
@mcp.tool()
async def word_Selection_SortByHeadings(this_Selection_wordObjId: str, SortFieldType, SortOrder, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID):
	"""This tool calls the SortByHeadings methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		SortFieldType: the SortFieldType as VT_VARIANT
		SortOrder: the SortOrder as VT_VARIANT
		CaseSensitive: the CaseSensitive as VT_VARIANT
		BidiSort: the BidiSort as VT_VARIANT
		IgnoreThe: the IgnoreThe as VT_VARIANT
		IgnoreKashida: the IgnoreKashida as VT_VARIANT
		IgnoreDiacritics: the IgnoreDiacritics as VT_VARIANT
		IgnoreHe: the IgnoreHe as VT_VARIANT
		LanguageID: the LanguageID as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	SortFieldType = tryParseString(SortFieldType)
	SortOrder = tryParseString(SortOrder)
	CaseSensitive = tryParseString(CaseSensitive)
	BidiSort = tryParseString(BidiSort)
	IgnoreThe = tryParseString(IgnoreThe)
	IgnoreKashida = tryParseString(IgnoreKashida)
	IgnoreDiacritics = tryParseString(IgnoreDiacritics)
	IgnoreHe = tryParseString(IgnoreHe)
	LanguageID = tryParseString(LanguageID)
	this_Selection.SortByHeadings(SortFieldType, SortOrder, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID)


# Tool: 529
@mcp.tool()
async def word_Selection_ExportAsFixedFormat2(this_Selection_wordObjId: str, OutputFileName: str, ExportFormat: int, OpenAfterExport: bool, OptimizeFor: int, ExportCurrentPage: bool, Item: int, IncludeDocProps: bool, KeepIRM: bool, CreateBookmarks: int, DocStructureTags: bool, BitmapMissingFonts: bool, UseISO19005_1: bool, OptimizeForImageQuality: bool, FixedFormatExtClassPtr):
	"""This tool calls the ExportAsFixedFormat2 methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		OutputFileName: the OutputFileName as str
		ExportFormat: the ExportFormat as WdExportFormat
		OpenAfterExport: the OpenAfterExport as bool
		OptimizeFor: the OptimizeFor as WdExportOptimizeFor
		ExportCurrentPage: the ExportCurrentPage as bool
		Item: the Item as WdExportItem
		IncludeDocProps: the IncludeDocProps as bool
		KeepIRM: the KeepIRM as bool
		CreateBookmarks: the CreateBookmarks as WdExportCreateBookmarks
		DocStructureTags: the DocStructureTags as bool
		BitmapMissingFonts: the BitmapMissingFonts as bool
		UseISO19005_1: the UseISO19005_1 as bool
		OptimizeForImageQuality: the OptimizeForImageQuality as bool
		FixedFormatExtClassPtr: the FixedFormatExtClassPtr as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	FixedFormatExtClassPtr = tryParseString(FixedFormatExtClassPtr)
	this_Selection.ExportAsFixedFormat2(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, OptimizeForImageQuality, FixedFormatExtClassPtr)


# Tool: 530
@mcp.tool()
async def word_Selection_ExportAsFixedFormat3(this_Selection_wordObjId: str, OutputFileName: str, ExportFormat: int, OpenAfterExport: bool, OptimizeFor: int, ExportCurrentPage: bool, Item: int, IncludeDocProps: bool, KeepIRM: bool, CreateBookmarks: int, DocStructureTags: bool, BitmapMissingFonts: bool, UseISO19005_1: bool, OptimizeForImageQuality: bool, ImproveExportTagging: bool, FixedFormatExtClassPtr):
	"""This tool calls the ExportAsFixedFormat3 methodon an Selection object. Pass the __WordObjectId of Selection of the object you want to call the method on as the first parameter
	
	Parameters:
		OutputFileName: the OutputFileName as str
		ExportFormat: the ExportFormat as WdExportFormat
		OpenAfterExport: the OpenAfterExport as bool
		OptimizeFor: the OptimizeFor as WdExportOptimizeFor
		ExportCurrentPage: the ExportCurrentPage as bool
		Item: the Item as WdExportItem
		IncludeDocProps: the IncludeDocProps as bool
		KeepIRM: the KeepIRM as bool
		CreateBookmarks: the CreateBookmarks as WdExportCreateBookmarks
		DocStructureTags: the DocStructureTags as bool
		BitmapMissingFonts: the BitmapMissingFonts as bool
		UseISO19005_1: the UseISO19005_1 as bool
		OptimizeForImageQuality: the OptimizeForImageQuality as bool
		ImproveExportTagging: the ImproveExportTagging as bool
		FixedFormatExtClassPtr: the FixedFormatExtClassPtr as VT_VARIANT
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	FixedFormatExtClassPtr = tryParseString(FixedFormatExtClassPtr)
	this_Selection.ExportAsFixedFormat3(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, OptimizeForImageQuality, ImproveExportTagging, FixedFormatExtClassPtr)


# Tool: 531
@mcp.tool()
async def word_Selection_get_Property(this_Selection_wordObjId: str, propertyName: str):
	"""Gets properties of Selection
	
	propertyName: Name of the property. Can be one of ...
		Text, FormattedText, Start, End, Font, Type, StoryType, Style, Tables, Words, Sentences, Characters, Footnotes, Endnotes, Comments, Cells, Sections, Paragraphs, Borders, Shading, Fields, FormFields, Frames, ParagraphFormat, PageSetup, Bookmarks, StoryLength, LanguageID, LanguageIDFarEast, LanguageIDOther, Hyperlinks, Columns, Rows, HeaderFooter, IsEndOfRowMark, BookmarkID, PreviousBookmarkID, Find, Range, Flags, Active, StartIsActive, IPAtEndOfLine, ExtendMode, ColumnSelectMode, Orientation, InlineShapes, Document, ShapeRange, NoProofing, TopLevelTables, LanguageDetected, FitTextWidth, HTMLDivisions, SmartTags, ChildShapeRange, HasChildShapeRange, FootnoteOptions, EndnoteOptions, XMLNodes, XMLParentNode, Editors, EnhMetaFileBits, OMaths, WordOpenXML, ContentControls, ParentContentControl
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	
	EnsureWord()
	if (propertyName == "Text"):
		retVal = this_Selection.Text
		return retVal
	if (propertyName == "FormattedText"):
		retVal = this_Selection.FormattedText
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_Start = retVal.Start
		except:
			local_Start = None
		try:
			local_End = retVal.End
		except:
			local_End = None
		try:
			local_StoryType = retVal.StoryType
		except:
			local_StoryType = None
		try:
			local_Bold = retVal.Bold
		except:
			local_Bold = None
		try:
			local_Italic = retVal.Italic
		except:
			local_Italic = None
		try:
			local_Underline = retVal.Underline
		except:
			local_Underline = None
		try:
			local_EmphasisMark = retVal.EmphasisMark
		except:
			local_EmphasisMark = None
		try:
			local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
		except:
			local_DisableCharacterSpaceGrid = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_StoryLength = retVal.StoryLength
		except:
			local_StoryLength = None
		try:
			local_LanguageID = retVal.LanguageID
		except:
			local_LanguageID = None
		try:
			local_GrammarChecked = retVal.GrammarChecked
		except:
			local_GrammarChecked = None
		try:
			local_SpellingChecked = retVal.SpellingChecked
		except:
			local_SpellingChecked = None
		try:
			local_HighlightColorIndex = retVal.HighlightColorIndex
		except:
			local_HighlightColorIndex = None
		try:
			local_CanEdit = retVal.CanEdit
		except:
			local_CanEdit = None
		try:
			local_CanPaste = retVal.CanPaste
		except:
			local_CanPaste = None
		try:
			local_IsEndOfRowMark = retVal.IsEndOfRowMark
		except:
			local_IsEndOfRowMark = None
		try:
			local_BookmarkID = retVal.BookmarkID
		except:
			local_BookmarkID = None
		try:
			local_PreviousBookmarkID = retVal.PreviousBookmarkID
		except:
			local_PreviousBookmarkID = None
		try:
			local_Case = retVal.Case
		except:
			local_Case = None
		try:
			local_Information = retVal.Information
		except:
			local_Information = None
		try:
			local_Orientation = retVal.Orientation
		except:
			local_Orientation = None
		try:
			local_LanguageIDFarEast = retVal.LanguageIDFarEast
		except:
			local_LanguageIDFarEast = None
		try:
			local_LanguageIDOther = retVal.LanguageIDOther
		except:
			local_LanguageIDOther = None
		try:
			local_LanguageDetected = retVal.LanguageDetected
		except:
			local_LanguageDetected = None
		try:
			local_FitTextWidth = retVal.FitTextWidth
		except:
			local_FitTextWidth = None
		try:
			local_HorizontalInVertical = retVal.HorizontalInVertical
		except:
			local_HorizontalInVertical = None
		try:
			local_TwoLinesInOne = retVal.TwoLinesInOne
		except:
			local_TwoLinesInOne = None
		try:
			local_CombineCharacters = retVal.CombineCharacters
		except:
			local_CombineCharacters = None
		try:
			local_NoProofing = retVal.NoProofing
		except:
			local_NoProofing = None
		try:
			local_CharacterWidth = retVal.CharacterWidth
		except:
			local_CharacterWidth = None
		try:
			local_Kana = retVal.Kana
		except:
			local_Kana = None
		try:
			local_BoldBi = retVal.BoldBi
		except:
			local_BoldBi = None
		try:
			local_ItalicBi = retVal.ItalicBi
		except:
			local_ItalicBi = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_ShowAll = retVal.ShowAll
		except:
			local_ShowAll = None
		try:
			local_CharacterStyle = retVal.CharacterStyle
		except:
			local_CharacterStyle = None
		try:
			local_ParagraphStyle = retVal.ParagraphStyle
		except:
			local_ParagraphStyle = None
		try:
			local_ListStyle = retVal.ListStyle
		except:
			local_ListStyle = None
		try:
			local_TableStyle = retVal.TableStyle
		except:
			local_TableStyle = None
		try:
			local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
		except:
			local_TextVisibleOnScreen = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }
	if (propertyName == "Start"):
		retVal = this_Selection.Start
		return retVal
	if (propertyName == "End"):
		retVal = this_Selection.End
		return retVal
	if (propertyName == "Font"):
		retVal = this_Selection.Font
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Font"}
	if (propertyName == "Type"):
		retVal = this_Selection.Type
		return retVal
	if (propertyName == "StoryType"):
		retVal = this_Selection.StoryType
		return retVal
	if (propertyName == "Style"):
		retVal = this_Selection.Style
		return retVal
	if (propertyName == "Tables"):
		retVal = this_Selection.Tables
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Tables", "Count": local_Count, "NestingLevel": local_NestingLevel, }
	if (propertyName == "Words"):
		retVal = this_Selection.Words
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Words", "Count": local_Count, }
	if (propertyName == "Sentences"):
		retVal = this_Selection.Sentences
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Sentences", "Count": local_Count, }
	if (propertyName == "Characters"):
		retVal = this_Selection.Characters
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Characters", "Count": local_Count, }
	if (propertyName == "Footnotes"):
		retVal = this_Selection.Footnotes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Location = retVal.Location
		except:
			local_Location = None
		try:
			local_NumberStyle = retVal.NumberStyle
		except:
			local_NumberStyle = None
		try:
			local_StartingNumber = retVal.StartingNumber
		except:
			local_StartingNumber = None
		try:
			local_NumberingRule = retVal.NumberingRule
		except:
			local_NumberingRule = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Footnotes", "Count": local_Count, "Location": local_Location, "NumberStyle": local_NumberStyle, "StartingNumber": local_StartingNumber, "NumberingRule": local_NumberingRule, }
	if (propertyName == "Endnotes"):
		retVal = this_Selection.Endnotes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Location = retVal.Location
		except:
			local_Location = None
		try:
			local_NumberStyle = retVal.NumberStyle
		except:
			local_NumberStyle = None
		try:
			local_StartingNumber = retVal.StartingNumber
		except:
			local_StartingNumber = None
		try:
			local_NumberingRule = retVal.NumberingRule
		except:
			local_NumberingRule = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Endnotes", "Count": local_Count, "Location": local_Location, "NumberStyle": local_NumberStyle, "StartingNumber": local_StartingNumber, "NumberingRule": local_NumberingRule, }
	if (propertyName == "Comments"):
		retVal = this_Selection.Comments
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_ShowBy = retVal.ShowBy
		except:
			local_ShowBy = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Comments", "Count": local_Count, "ShowBy": local_ShowBy, }
	if (propertyName == "Cells"):
		retVal = this_Selection.Cells
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_VerticalAlignment = retVal.VerticalAlignment
		except:
			local_VerticalAlignment = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Cells", "Count": local_Count, "Width": local_Width, "Height": local_Height, "HeightRule": local_HeightRule, "VerticalAlignment": local_VerticalAlignment, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "Sections"):
		retVal = this_Selection.Sections
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Sections", "Count": local_Count, }
	if (propertyName == "Paragraphs"):
		retVal = this_Selection.Paragraphs
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_Alignment = retVal.Alignment
		except:
			local_Alignment = None
		try:
			local_KeepTogether = retVal.KeepTogether
		except:
			local_KeepTogether = None
		try:
			local_KeepWithNext = retVal.KeepWithNext
		except:
			local_KeepWithNext = None
		try:
			local_PageBreakBefore = retVal.PageBreakBefore
		except:
			local_PageBreakBefore = None
		try:
			local_NoLineNumber = retVal.NoLineNumber
		except:
			local_NoLineNumber = None
		try:
			local_RightIndent = retVal.RightIndent
		except:
			local_RightIndent = None
		try:
			local_LeftIndent = retVal.LeftIndent
		except:
			local_LeftIndent = None
		try:
			local_FirstLineIndent = retVal.FirstLineIndent
		except:
			local_FirstLineIndent = None
		try:
			local_LineSpacing = retVal.LineSpacing
		except:
			local_LineSpacing = None
		try:
			local_LineSpacingRule = retVal.LineSpacingRule
		except:
			local_LineSpacingRule = None
		try:
			local_SpaceBefore = retVal.SpaceBefore
		except:
			local_SpaceBefore = None
		try:
			local_SpaceAfter = retVal.SpaceAfter
		except:
			local_SpaceAfter = None
		try:
			local_Hyphenation = retVal.Hyphenation
		except:
			local_Hyphenation = None
		try:
			local_WidowControl = retVal.WidowControl
		except:
			local_WidowControl = None
		try:
			local_FarEastLineBreakControl = retVal.FarEastLineBreakControl
		except:
			local_FarEastLineBreakControl = None
		try:
			local_WordWrap = retVal.WordWrap
		except:
			local_WordWrap = None
		try:
			local_HangingPunctuation = retVal.HangingPunctuation
		except:
			local_HangingPunctuation = None
		try:
			local_HalfWidthPunctuationOnTopOfLine = retVal.HalfWidthPunctuationOnTopOfLine
		except:
			local_HalfWidthPunctuationOnTopOfLine = None
		try:
			local_AddSpaceBetweenFarEastAndAlpha = retVal.AddSpaceBetweenFarEastAndAlpha
		except:
			local_AddSpaceBetweenFarEastAndAlpha = None
		try:
			local_AddSpaceBetweenFarEastAndDigit = retVal.AddSpaceBetweenFarEastAndDigit
		except:
			local_AddSpaceBetweenFarEastAndDigit = None
		try:
			local_BaseLineAlignment = retVal.BaseLineAlignment
		except:
			local_BaseLineAlignment = None
		try:
			local_AutoAdjustRightIndent = retVal.AutoAdjustRightIndent
		except:
			local_AutoAdjustRightIndent = None
		try:
			local_DisableLineHeightGrid = retVal.DisableLineHeightGrid
		except:
			local_DisableLineHeightGrid = None
		try:
			local_OutlineLevel = retVal.OutlineLevel
		except:
			local_OutlineLevel = None
		try:
			local_CharacterUnitRightIndent = retVal.CharacterUnitRightIndent
		except:
			local_CharacterUnitRightIndent = None
		try:
			local_CharacterUnitLeftIndent = retVal.CharacterUnitLeftIndent
		except:
			local_CharacterUnitLeftIndent = None
		try:
			local_CharacterUnitFirstLineIndent = retVal.CharacterUnitFirstLineIndent
		except:
			local_CharacterUnitFirstLineIndent = None
		try:
			local_LineUnitBefore = retVal.LineUnitBefore
		except:
			local_LineUnitBefore = None
		try:
			local_LineUnitAfter = retVal.LineUnitAfter
		except:
			local_LineUnitAfter = None
		try:
			local_ReadingOrder = retVal.ReadingOrder
		except:
			local_ReadingOrder = None
		try:
			local_SpaceBeforeAuto = retVal.SpaceBeforeAuto
		except:
			local_SpaceBeforeAuto = None
		try:
			local_SpaceAfterAuto = retVal.SpaceAfterAuto
		except:
			local_SpaceAfterAuto = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Paragraphs", "Count": local_Count, "Style": local_Style, "Alignment": local_Alignment, "KeepTogether": local_KeepTogether, "KeepWithNext": local_KeepWithNext, "PageBreakBefore": local_PageBreakBefore, "NoLineNumber": local_NoLineNumber, "RightIndent": local_RightIndent, "LeftIndent": local_LeftIndent, "FirstLineIndent": local_FirstLineIndent, "LineSpacing": local_LineSpacing, "LineSpacingRule": local_LineSpacingRule, "SpaceBefore": local_SpaceBefore, "SpaceAfter": local_SpaceAfter, "Hyphenation": local_Hyphenation, "WidowControl": local_WidowControl, "FarEastLineBreakControl": local_FarEastLineBreakControl, "WordWrap": local_WordWrap, "HangingPunctuation": local_HangingPunctuation, "HalfWidthPunctuationOnTopOfLine": local_HalfWidthPunctuationOnTopOfLine, "AddSpaceBetweenFarEastAndAlpha": local_AddSpaceBetweenFarEastAndAlpha, "AddSpaceBetweenFarEastAndDigit": local_AddSpaceBetweenFarEastAndDigit, "BaseLineAlignment": local_BaseLineAlignment, "AutoAdjustRightIndent": local_AutoAdjustRightIndent, "DisableLineHeightGrid": local_DisableLineHeightGrid, "OutlineLevel": local_OutlineLevel, "CharacterUnitRightIndent": local_CharacterUnitRightIndent, "CharacterUnitLeftIndent": local_CharacterUnitLeftIndent, "CharacterUnitFirstLineIndent": local_CharacterUnitFirstLineIndent, "LineUnitBefore": local_LineUnitBefore, "LineUnitAfter": local_LineUnitAfter, "ReadingOrder": local_ReadingOrder, "SpaceBeforeAuto": local_SpaceBeforeAuto, "SpaceAfterAuto": local_SpaceAfterAuto, }
	if (propertyName == "Borders"):
		retVal = this_Selection.Borders
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Enable = retVal.Enable
		except:
			local_Enable = None
		try:
			local_DistanceFromTop = retVal.DistanceFromTop
		except:
			local_DistanceFromTop = None
		try:
			local_Shadow = retVal.Shadow
		except:
			local_Shadow = None
		try:
			local_InsideLineStyle = retVal.InsideLineStyle
		except:
			local_InsideLineStyle = None
		try:
			local_OutsideLineStyle = retVal.OutsideLineStyle
		except:
			local_OutsideLineStyle = None
		try:
			local_InsideLineWidth = retVal.InsideLineWidth
		except:
			local_InsideLineWidth = None
		try:
			local_OutsideLineWidth = retVal.OutsideLineWidth
		except:
			local_OutsideLineWidth = None
		try:
			local_InsideColorIndex = retVal.InsideColorIndex
		except:
			local_InsideColorIndex = None
		try:
			local_OutsideColorIndex = retVal.OutsideColorIndex
		except:
			local_OutsideColorIndex = None
		try:
			local_DistanceFromLeft = retVal.DistanceFromLeft
		except:
			local_DistanceFromLeft = None
		try:
			local_DistanceFromBottom = retVal.DistanceFromBottom
		except:
			local_DistanceFromBottom = None
		try:
			local_DistanceFromRight = retVal.DistanceFromRight
		except:
			local_DistanceFromRight = None
		try:
			local_AlwaysInFront = retVal.AlwaysInFront
		except:
			local_AlwaysInFront = None
		try:
			local_SurroundHeader = retVal.SurroundHeader
		except:
			local_SurroundHeader = None
		try:
			local_SurroundFooter = retVal.SurroundFooter
		except:
			local_SurroundFooter = None
		try:
			local_JoinBorders = retVal.JoinBorders
		except:
			local_JoinBorders = None
		try:
			local_HasHorizontal = retVal.HasHorizontal
		except:
			local_HasHorizontal = None
		try:
			local_HasVertical = retVal.HasVertical
		except:
			local_HasVertical = None
		try:
			local_DistanceFrom = retVal.DistanceFrom
		except:
			local_DistanceFrom = None
		try:
			local_EnableFirstPageInSection = retVal.EnableFirstPageInSection
		except:
			local_EnableFirstPageInSection = None
		try:
			local_EnableOtherPagesInSection = retVal.EnableOtherPagesInSection
		except:
			local_EnableOtherPagesInSection = None
		try:
			local_InsideColor = retVal.InsideColor
		except:
			local_InsideColor = None
		try:
			local_OutsideColor = retVal.OutsideColor
		except:
			local_OutsideColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Borders", "Count": local_Count, "Enable": local_Enable, "DistanceFromTop": local_DistanceFromTop, "Shadow": local_Shadow, "InsideLineStyle": local_InsideLineStyle, "OutsideLineStyle": local_OutsideLineStyle, "InsideLineWidth": local_InsideLineWidth, "OutsideLineWidth": local_OutsideLineWidth, "InsideColorIndex": local_InsideColorIndex, "OutsideColorIndex": local_OutsideColorIndex, "DistanceFromLeft": local_DistanceFromLeft, "DistanceFromBottom": local_DistanceFromBottom, "DistanceFromRight": local_DistanceFromRight, "AlwaysInFront": local_AlwaysInFront, "SurroundHeader": local_SurroundHeader, "SurroundFooter": local_SurroundFooter, "JoinBorders": local_JoinBorders, "HasHorizontal": local_HasHorizontal, "HasVertical": local_HasVertical, "DistanceFrom": local_DistanceFrom, "EnableFirstPageInSection": local_EnableFirstPageInSection, "EnableOtherPagesInSection": local_EnableOtherPagesInSection, "InsideColor": local_InsideColor, "OutsideColor": local_OutsideColor, }
	if (propertyName == "Shading"):
		retVal = this_Selection.Shading
		try:
			local_ForegroundPatternColorIndex = retVal.ForegroundPatternColorIndex
		except:
			local_ForegroundPatternColorIndex = None
		try:
			local_BackgroundPatternColorIndex = retVal.BackgroundPatternColorIndex
		except:
			local_BackgroundPatternColorIndex = None
		try:
			local_Texture = retVal.Texture
		except:
			local_Texture = None
		try:
			local_ForegroundPatternColor = retVal.ForegroundPatternColor
		except:
			local_ForegroundPatternColor = None
		try:
			local_BackgroundPatternColor = retVal.BackgroundPatternColor
		except:
			local_BackgroundPatternColor = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Shading", "ForegroundPatternColorIndex": local_ForegroundPatternColorIndex, "BackgroundPatternColorIndex": local_BackgroundPatternColorIndex, "Texture": local_Texture, "ForegroundPatternColor": local_ForegroundPatternColor, "BackgroundPatternColor": local_BackgroundPatternColor, }
	if (propertyName == "Fields"):
		retVal = this_Selection.Fields
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Locked = retVal.Locked
		except:
			local_Locked = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Fields", "Count": local_Count, "Locked": local_Locked, }
	if (propertyName == "FormFields"):
		retVal = this_Selection.FormFields
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Shaded = retVal.Shaded
		except:
			local_Shaded = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "FormFields", "Count": local_Count, "Shaded": local_Shaded, }
	if (propertyName == "Frames"):
		retVal = this_Selection.Frames
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Frames", "Count": local_Count, }
	if (propertyName == "ParagraphFormat"):
		retVal = this_Selection.ParagraphFormat
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ParagraphFormat"}
	if (propertyName == "PageSetup"):
		retVal = this_Selection.PageSetup
		try:
			local_TopMargin = retVal.TopMargin
		except:
			local_TopMargin = None
		try:
			local_BottomMargin = retVal.BottomMargin
		except:
			local_BottomMargin = None
		try:
			local_LeftMargin = retVal.LeftMargin
		except:
			local_LeftMargin = None
		try:
			local_RightMargin = retVal.RightMargin
		except:
			local_RightMargin = None
		try:
			local_Gutter = retVal.Gutter
		except:
			local_Gutter = None
		try:
			local_PageWidth = retVal.PageWidth
		except:
			local_PageWidth = None
		try:
			local_PageHeight = retVal.PageHeight
		except:
			local_PageHeight = None
		try:
			local_Orientation = retVal.Orientation
		except:
			local_Orientation = None
		try:
			local_FirstPageTray = retVal.FirstPageTray
		except:
			local_FirstPageTray = None
		try:
			local_OtherPagesTray = retVal.OtherPagesTray
		except:
			local_OtherPagesTray = None
		try:
			local_VerticalAlignment = retVal.VerticalAlignment
		except:
			local_VerticalAlignment = None
		try:
			local_MirrorMargins = retVal.MirrorMargins
		except:
			local_MirrorMargins = None
		try:
			local_HeaderDistance = retVal.HeaderDistance
		except:
			local_HeaderDistance = None
		try:
			local_FooterDistance = retVal.FooterDistance
		except:
			local_FooterDistance = None
		try:
			local_SectionStart = retVal.SectionStart
		except:
			local_SectionStart = None
		try:
			local_OddAndEvenPagesHeaderFooter = retVal.OddAndEvenPagesHeaderFooter
		except:
			local_OddAndEvenPagesHeaderFooter = None
		try:
			local_DifferentFirstPageHeaderFooter = retVal.DifferentFirstPageHeaderFooter
		except:
			local_DifferentFirstPageHeaderFooter = None
		try:
			local_SuppressEndnotes = retVal.SuppressEndnotes
		except:
			local_SuppressEndnotes = None
		try:
			local_PaperSize = retVal.PaperSize
		except:
			local_PaperSize = None
		try:
			local_TwoPagesOnOne = retVal.TwoPagesOnOne
		except:
			local_TwoPagesOnOne = None
		try:
			local_GutterOnTop = retVal.GutterOnTop
		except:
			local_GutterOnTop = None
		try:
			local_CharsLine = retVal.CharsLine
		except:
			local_CharsLine = None
		try:
			local_LinesPage = retVal.LinesPage
		except:
			local_LinesPage = None
		try:
			local_ShowGrid = retVal.ShowGrid
		except:
			local_ShowGrid = None
		try:
			local_GutterStyle = retVal.GutterStyle
		except:
			local_GutterStyle = None
		try:
			local_SectionDirection = retVal.SectionDirection
		except:
			local_SectionDirection = None
		try:
			local_LayoutMode = retVal.LayoutMode
		except:
			local_LayoutMode = None
		try:
			local_GutterPos = retVal.GutterPos
		except:
			local_GutterPos = None
		try:
			local_BookFoldPrinting = retVal.BookFoldPrinting
		except:
			local_BookFoldPrinting = None
		try:
			local_BookFoldRevPrinting = retVal.BookFoldRevPrinting
		except:
			local_BookFoldRevPrinting = None
		try:
			local_BookFoldPrintingSheets = retVal.BookFoldPrintingSheets
		except:
			local_BookFoldPrintingSheets = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "PageSetup", "TopMargin": local_TopMargin, "BottomMargin": local_BottomMargin, "LeftMargin": local_LeftMargin, "RightMargin": local_RightMargin, "Gutter": local_Gutter, "PageWidth": local_PageWidth, "PageHeight": local_PageHeight, "Orientation": local_Orientation, "FirstPageTray": local_FirstPageTray, "OtherPagesTray": local_OtherPagesTray, "VerticalAlignment": local_VerticalAlignment, "MirrorMargins": local_MirrorMargins, "HeaderDistance": local_HeaderDistance, "FooterDistance": local_FooterDistance, "SectionStart": local_SectionStart, "OddAndEvenPagesHeaderFooter": local_OddAndEvenPagesHeaderFooter, "DifferentFirstPageHeaderFooter": local_DifferentFirstPageHeaderFooter, "SuppressEndnotes": local_SuppressEndnotes, "PaperSize": local_PaperSize, "TwoPagesOnOne": local_TwoPagesOnOne, "GutterOnTop": local_GutterOnTop, "CharsLine": local_CharsLine, "LinesPage": local_LinesPage, "ShowGrid": local_ShowGrid, "GutterStyle": local_GutterStyle, "SectionDirection": local_SectionDirection, "LayoutMode": local_LayoutMode, "GutterPos": local_GutterPos, "BookFoldPrinting": local_BookFoldPrinting, "BookFoldRevPrinting": local_BookFoldRevPrinting, "BookFoldPrintingSheets": local_BookFoldPrintingSheets, }
	if (propertyName == "Bookmarks"):
		retVal = this_Selection.Bookmarks
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_DefaultSorting = retVal.DefaultSorting
		except:
			local_DefaultSorting = None
		try:
			local_ShowHidden = retVal.ShowHidden
		except:
			local_ShowHidden = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Bookmarks", "Count": local_Count, "DefaultSorting": local_DefaultSorting, "ShowHidden": local_ShowHidden, }
	if (propertyName == "StoryLength"):
		retVal = this_Selection.StoryLength
		return retVal
	if (propertyName == "LanguageID"):
		retVal = this_Selection.LanguageID
		return retVal
	if (propertyName == "LanguageIDFarEast"):
		retVal = this_Selection.LanguageIDFarEast
		return retVal
	if (propertyName == "LanguageIDOther"):
		retVal = this_Selection.LanguageIDOther
		return retVal
	if (propertyName == "Hyperlinks"):
		retVal = this_Selection.Hyperlinks
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Hyperlinks", "Count": local_Count, }
	if (propertyName == "Columns"):
		retVal = this_Selection.Columns
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_PreferredWidth = retVal.PreferredWidth
		except:
			local_PreferredWidth = None
		try:
			local_PreferredWidthType = retVal.PreferredWidthType
		except:
			local_PreferredWidthType = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Columns", "Count": local_Count, "Width": local_Width, "NestingLevel": local_NestingLevel, "PreferredWidth": local_PreferredWidth, "PreferredWidthType": local_PreferredWidthType, }
	if (propertyName == "Rows"):
		retVal = this_Selection.Rows
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_AllowBreakAcrossPages = retVal.AllowBreakAcrossPages
		except:
			local_AllowBreakAcrossPages = None
		try:
			local_Alignment = retVal.Alignment
		except:
			local_Alignment = None
		try:
			local_HeadingFormat = retVal.HeadingFormat
		except:
			local_HeadingFormat = None
		try:
			local_SpaceBetweenColumns = retVal.SpaceBetweenColumns
		except:
			local_SpaceBetweenColumns = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HeightRule = retVal.HeightRule
		except:
			local_HeightRule = None
		try:
			local_LeftIndent = retVal.LeftIndent
		except:
			local_LeftIndent = None
		try:
			local_WrapAroundText = retVal.WrapAroundText
		except:
			local_WrapAroundText = None
		try:
			local_DistanceTop = retVal.DistanceTop
		except:
			local_DistanceTop = None
		try:
			local_DistanceBottom = retVal.DistanceBottom
		except:
			local_DistanceBottom = None
		try:
			local_DistanceLeft = retVal.DistanceLeft
		except:
			local_DistanceLeft = None
		try:
			local_DistanceRight = retVal.DistanceRight
		except:
			local_DistanceRight = None
		try:
			local_HorizontalPosition = retVal.HorizontalPosition
		except:
			local_HorizontalPosition = None
		try:
			local_VerticalPosition = retVal.VerticalPosition
		except:
			local_VerticalPosition = None
		try:
			local_RelativeHorizontalPosition = retVal.RelativeHorizontalPosition
		except:
			local_RelativeHorizontalPosition = None
		try:
			local_RelativeVerticalPosition = retVal.RelativeVerticalPosition
		except:
			local_RelativeVerticalPosition = None
		try:
			local_AllowOverlap = retVal.AllowOverlap
		except:
			local_AllowOverlap = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		try:
			local_TableDirection = retVal.TableDirection
		except:
			local_TableDirection = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Rows", "Count": local_Count, "AllowBreakAcrossPages": local_AllowBreakAcrossPages, "Alignment": local_Alignment, "HeadingFormat": local_HeadingFormat, "SpaceBetweenColumns": local_SpaceBetweenColumns, "Height": local_Height, "HeightRule": local_HeightRule, "LeftIndent": local_LeftIndent, "WrapAroundText": local_WrapAroundText, "DistanceTop": local_DistanceTop, "DistanceBottom": local_DistanceBottom, "DistanceLeft": local_DistanceLeft, "DistanceRight": local_DistanceRight, "HorizontalPosition": local_HorizontalPosition, "VerticalPosition": local_VerticalPosition, "RelativeHorizontalPosition": local_RelativeHorizontalPosition, "RelativeVerticalPosition": local_RelativeVerticalPosition, "AllowOverlap": local_AllowOverlap, "NestingLevel": local_NestingLevel, "TableDirection": local_TableDirection, }
	if (propertyName == "HeaderFooter"):
		retVal = this_Selection.HeaderFooter
		try:
			local_Index = retVal.Index
		except:
			local_Index = None
		try:
			local_IsHeader = retVal.IsHeader
		except:
			local_IsHeader = None
		try:
			local_Exists = retVal.Exists
		except:
			local_Exists = None
		try:
			local_LinkToPrevious = retVal.LinkToPrevious
		except:
			local_LinkToPrevious = None
		try:
			local_IsEmpty = retVal.IsEmpty
		except:
			local_IsEmpty = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "HeaderFooter", "Index": local_Index, "IsHeader": local_IsHeader, "Exists": local_Exists, "LinkToPrevious": local_LinkToPrevious, "IsEmpty": local_IsEmpty, }
	if (propertyName == "IsEndOfRowMark"):
		retVal = this_Selection.IsEndOfRowMark
		return retVal
	if (propertyName == "BookmarkID"):
		retVal = this_Selection.BookmarkID
		return retVal
	if (propertyName == "PreviousBookmarkID"):
		retVal = this_Selection.PreviousBookmarkID
		return retVal
	if (propertyName == "Find"):
		retVal = this_Selection.Find
		try:
			local_Forward = retVal.Forward
		except:
			local_Forward = None
		try:
			local_Found = retVal.Found
		except:
			local_Found = None
		try:
			local_MatchAllWordForms = retVal.MatchAllWordForms
		except:
			local_MatchAllWordForms = None
		try:
			local_MatchCase = retVal.MatchCase
		except:
			local_MatchCase = None
		try:
			local_MatchWildcards = retVal.MatchWildcards
		except:
			local_MatchWildcards = None
		try:
			local_MatchSoundsLike = retVal.MatchSoundsLike
		except:
			local_MatchSoundsLike = None
		try:
			local_MatchWholeWord = retVal.MatchWholeWord
		except:
			local_MatchWholeWord = None
		try:
			local_MatchFuzzy = retVal.MatchFuzzy
		except:
			local_MatchFuzzy = None
		try:
			local_MatchByte = retVal.MatchByte
		except:
			local_MatchByte = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_LanguageID = retVal.LanguageID
		except:
			local_LanguageID = None
		try:
			local_Highlight = retVal.Highlight
		except:
			local_Highlight = None
		try:
			local_Wrap = retVal.Wrap
		except:
			local_Wrap = None
		try:
			local_Format = retVal.Format
		except:
			local_Format = None
		try:
			local_LanguageIDFarEast = retVal.LanguageIDFarEast
		except:
			local_LanguageIDFarEast = None
		try:
			local_LanguageIDOther = retVal.LanguageIDOther
		except:
			local_LanguageIDOther = None
		try:
			local_CorrectHangulEndings = retVal.CorrectHangulEndings
		except:
			local_CorrectHangulEndings = None
		try:
			local_NoProofing = retVal.NoProofing
		except:
			local_NoProofing = None
		try:
			local_MatchKashida = retVal.MatchKashida
		except:
			local_MatchKashida = None
		try:
			local_MatchDiacritics = retVal.MatchDiacritics
		except:
			local_MatchDiacritics = None
		try:
			local_MatchAlefHamza = retVal.MatchAlefHamza
		except:
			local_MatchAlefHamza = None
		try:
			local_MatchControl = retVal.MatchControl
		except:
			local_MatchControl = None
		try:
			local_MatchPhrase = retVal.MatchPhrase
		except:
			local_MatchPhrase = None
		try:
			local_MatchPrefix = retVal.MatchPrefix
		except:
			local_MatchPrefix = None
		try:
			local_MatchSuffix = retVal.MatchSuffix
		except:
			local_MatchSuffix = None
		try:
			local_IgnoreSpace = retVal.IgnoreSpace
		except:
			local_IgnoreSpace = None
		try:
			local_IgnorePunct = retVal.IgnorePunct
		except:
			local_IgnorePunct = None
		try:
			local_HanjaPhoneticHangul = retVal.HanjaPhoneticHangul
		except:
			local_HanjaPhoneticHangul = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Find", "Forward": local_Forward, "Found": local_Found, "MatchAllWordForms": local_MatchAllWordForms, "MatchCase": local_MatchCase, "MatchWildcards": local_MatchWildcards, "MatchSoundsLike": local_MatchSoundsLike, "MatchWholeWord": local_MatchWholeWord, "MatchFuzzy": local_MatchFuzzy, "MatchByte": local_MatchByte, "Style": local_Style, "Text": local_Text, "LanguageID": local_LanguageID, "Highlight": local_Highlight, "Wrap": local_Wrap, "Format": local_Format, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "CorrectHangulEndings": local_CorrectHangulEndings, "NoProofing": local_NoProofing, "MatchKashida": local_MatchKashida, "MatchDiacritics": local_MatchDiacritics, "MatchAlefHamza": local_MatchAlefHamza, "MatchControl": local_MatchControl, "MatchPhrase": local_MatchPhrase, "MatchPrefix": local_MatchPrefix, "MatchSuffix": local_MatchSuffix, "IgnoreSpace": local_IgnoreSpace, "IgnorePunct": local_IgnorePunct, "HanjaPhoneticHangul": local_HanjaPhoneticHangul, }
	if (propertyName == "Range"):
		retVal = this_Selection.Range
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_Start = retVal.Start
		except:
			local_Start = None
		try:
			local_End = retVal.End
		except:
			local_End = None
		try:
			local_StoryType = retVal.StoryType
		except:
			local_StoryType = None
		try:
			local_Bold = retVal.Bold
		except:
			local_Bold = None
		try:
			local_Italic = retVal.Italic
		except:
			local_Italic = None
		try:
			local_Underline = retVal.Underline
		except:
			local_Underline = None
		try:
			local_EmphasisMark = retVal.EmphasisMark
		except:
			local_EmphasisMark = None
		try:
			local_DisableCharacterSpaceGrid = retVal.DisableCharacterSpaceGrid
		except:
			local_DisableCharacterSpaceGrid = None
		try:
			local_Style = retVal.Style
		except:
			local_Style = None
		try:
			local_StoryLength = retVal.StoryLength
		except:
			local_StoryLength = None
		try:
			local_LanguageID = retVal.LanguageID
		except:
			local_LanguageID = None
		try:
			local_GrammarChecked = retVal.GrammarChecked
		except:
			local_GrammarChecked = None
		try:
			local_SpellingChecked = retVal.SpellingChecked
		except:
			local_SpellingChecked = None
		try:
			local_HighlightColorIndex = retVal.HighlightColorIndex
		except:
			local_HighlightColorIndex = None
		try:
			local_CanEdit = retVal.CanEdit
		except:
			local_CanEdit = None
		try:
			local_CanPaste = retVal.CanPaste
		except:
			local_CanPaste = None
		try:
			local_IsEndOfRowMark = retVal.IsEndOfRowMark
		except:
			local_IsEndOfRowMark = None
		try:
			local_BookmarkID = retVal.BookmarkID
		except:
			local_BookmarkID = None
		try:
			local_PreviousBookmarkID = retVal.PreviousBookmarkID
		except:
			local_PreviousBookmarkID = None
		try:
			local_Case = retVal.Case
		except:
			local_Case = None
		try:
			local_Information = retVal.Information
		except:
			local_Information = None
		try:
			local_Orientation = retVal.Orientation
		except:
			local_Orientation = None
		try:
			local_LanguageIDFarEast = retVal.LanguageIDFarEast
		except:
			local_LanguageIDFarEast = None
		try:
			local_LanguageIDOther = retVal.LanguageIDOther
		except:
			local_LanguageIDOther = None
		try:
			local_LanguageDetected = retVal.LanguageDetected
		except:
			local_LanguageDetected = None
		try:
			local_FitTextWidth = retVal.FitTextWidth
		except:
			local_FitTextWidth = None
		try:
			local_HorizontalInVertical = retVal.HorizontalInVertical
		except:
			local_HorizontalInVertical = None
		try:
			local_TwoLinesInOne = retVal.TwoLinesInOne
		except:
			local_TwoLinesInOne = None
		try:
			local_CombineCharacters = retVal.CombineCharacters
		except:
			local_CombineCharacters = None
		try:
			local_NoProofing = retVal.NoProofing
		except:
			local_NoProofing = None
		try:
			local_CharacterWidth = retVal.CharacterWidth
		except:
			local_CharacterWidth = None
		try:
			local_Kana = retVal.Kana
		except:
			local_Kana = None
		try:
			local_BoldBi = retVal.BoldBi
		except:
			local_BoldBi = None
		try:
			local_ItalicBi = retVal.ItalicBi
		except:
			local_ItalicBi = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_ShowAll = retVal.ShowAll
		except:
			local_ShowAll = None
		try:
			local_CharacterStyle = retVal.CharacterStyle
		except:
			local_CharacterStyle = None
		try:
			local_ParagraphStyle = retVal.ParagraphStyle
		except:
			local_ParagraphStyle = None
		try:
			local_ListStyle = retVal.ListStyle
		except:
			local_ListStyle = None
		try:
			local_TableStyle = retVal.TableStyle
		except:
			local_TableStyle = None
		try:
			local_TextVisibleOnScreen = retVal.TextVisibleOnScreen
		except:
			local_TextVisibleOnScreen = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Range", "Text": local_Text, "Start": local_Start, "End": local_End, "StoryType": local_StoryType, "Bold": local_Bold, "Italic": local_Italic, "Underline": local_Underline, "EmphasisMark": local_EmphasisMark, "DisableCharacterSpaceGrid": local_DisableCharacterSpaceGrid, "Style": local_Style, "StoryLength": local_StoryLength, "LanguageID": local_LanguageID, "GrammarChecked": local_GrammarChecked, "SpellingChecked": local_SpellingChecked, "HighlightColorIndex": local_HighlightColorIndex, "CanEdit": local_CanEdit, "CanPaste": local_CanPaste, "IsEndOfRowMark": local_IsEndOfRowMark, "BookmarkID": local_BookmarkID, "PreviousBookmarkID": local_PreviousBookmarkID, "Case": local_Case, "Information": local_Information, "Orientation": local_Orientation, "LanguageIDFarEast": local_LanguageIDFarEast, "LanguageIDOther": local_LanguageIDOther, "LanguageDetected": local_LanguageDetected, "FitTextWidth": local_FitTextWidth, "HorizontalInVertical": local_HorizontalInVertical, "TwoLinesInOne": local_TwoLinesInOne, "CombineCharacters": local_CombineCharacters, "NoProofing": local_NoProofing, "CharacterWidth": local_CharacterWidth, "Kana": local_Kana, "BoldBi": local_BoldBi, "ItalicBi": local_ItalicBi, "ID": local_ID, "ShowAll": local_ShowAll, "CharacterStyle": local_CharacterStyle, "ParagraphStyle": local_ParagraphStyle, "ListStyle": local_ListStyle, "TableStyle": local_TableStyle, "TextVisibleOnScreen": local_TextVisibleOnScreen, }
	if (propertyName == "Flags"):
		retVal = this_Selection.Flags
		return retVal
	if (propertyName == "Active"):
		retVal = this_Selection.Active
		return retVal
	if (propertyName == "StartIsActive"):
		retVal = this_Selection.StartIsActive
		return retVal
	if (propertyName == "IPAtEndOfLine"):
		retVal = this_Selection.IPAtEndOfLine
		return retVal
	if (propertyName == "ExtendMode"):
		retVal = this_Selection.ExtendMode
		return retVal
	if (propertyName == "ColumnSelectMode"):
		retVal = this_Selection.ColumnSelectMode
		return retVal
	if (propertyName == "Orientation"):
		retVal = this_Selection.Orientation
		return retVal
	if (propertyName == "InlineShapes"):
		retVal = this_Selection.InlineShapes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "InlineShapes", "Count": local_Count, }
	if (propertyName == "Document"):
		retVal = this_Selection.Document
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Document"}
	if (propertyName == "ShapeRange"):
		retVal = this_Selection.ShapeRange
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_AutoShapeType = retVal.AutoShapeType
		except:
			local_AutoShapeType = None
		try:
			local_ConnectionSiteCount = retVal.ConnectionSiteCount
		except:
			local_ConnectionSiteCount = None
		try:
			local_Connector = retVal.Connector
		except:
			local_Connector = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HorizontalFlip = retVal.HorizontalFlip
		except:
			local_HorizontalFlip = None
		try:
			local_Left = retVal.Left
		except:
			local_Left = None
		try:
			local_LockAspectRatio = retVal.LockAspectRatio
		except:
			local_LockAspectRatio = None
		try:
			local_Name = retVal.Name
		except:
			local_Name = None
		try:
			local_Rotation = retVal.Rotation
		except:
			local_Rotation = None
		try:
			local_Top = retVal.Top
		except:
			local_Top = None
		try:
			local_Type = retVal.Type
		except:
			local_Type = None
		try:
			local_VerticalFlip = retVal.VerticalFlip
		except:
			local_VerticalFlip = None
		try:
			local_Vertices = retVal.Vertices
		except:
			local_Vertices = None
		try:
			local_Visible = retVal.Visible
		except:
			local_Visible = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_ZOrderPosition = retVal.ZOrderPosition
		except:
			local_ZOrderPosition = None
		try:
			local_RelativeHorizontalPosition = retVal.RelativeHorizontalPosition
		except:
			local_RelativeHorizontalPosition = None
		try:
			local_RelativeVerticalPosition = retVal.RelativeVerticalPosition
		except:
			local_RelativeVerticalPosition = None
		try:
			local_LockAnchor = retVal.LockAnchor
		except:
			local_LockAnchor = None
		try:
			local_AlternativeText = retVal.AlternativeText
		except:
			local_AlternativeText = None
		try:
			local_HasDiagram = retVal.HasDiagram
		except:
			local_HasDiagram = None
		try:
			local_HasDiagramNode = retVal.HasDiagramNode
		except:
			local_HasDiagramNode = None
		try:
			local_Child = retVal.Child
		except:
			local_Child = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_LayoutInCell = retVal.LayoutInCell
		except:
			local_LayoutInCell = None
		try:
			local_LeftRelative = retVal.LeftRelative
		except:
			local_LeftRelative = None
		try:
			local_TopRelative = retVal.TopRelative
		except:
			local_TopRelative = None
		try:
			local_WidthRelative = retVal.WidthRelative
		except:
			local_WidthRelative = None
		try:
			local_HeightRelative = retVal.HeightRelative
		except:
			local_HeightRelative = None
		try:
			local_RelativeHorizontalSize = retVal.RelativeHorizontalSize
		except:
			local_RelativeHorizontalSize = None
		try:
			local_RelativeVerticalSize = retVal.RelativeVerticalSize
		except:
			local_RelativeVerticalSize = None
		try:
			local_ShapeStyle = retVal.ShapeStyle
		except:
			local_ShapeStyle = None
		try:
			local_BackgroundStyle = retVal.BackgroundStyle
		except:
			local_BackgroundStyle = None
		try:
			local_Title = retVal.Title
		except:
			local_Title = None
		try:
			local_GraphicStyle = retVal.GraphicStyle
		except:
			local_GraphicStyle = None
		try:
			local_Decorative = retVal.Decorative
		except:
			local_Decorative = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ShapeRange", "Count": local_Count, "AutoShapeType": local_AutoShapeType, "ConnectionSiteCount": local_ConnectionSiteCount, "Connector": local_Connector, "Height": local_Height, "HorizontalFlip": local_HorizontalFlip, "Left": local_Left, "LockAspectRatio": local_LockAspectRatio, "Name": local_Name, "Rotation": local_Rotation, "Top": local_Top, "Type": local_Type, "VerticalFlip": local_VerticalFlip, "Vertices": local_Vertices, "Visible": local_Visible, "Width": local_Width, "ZOrderPosition": local_ZOrderPosition, "RelativeHorizontalPosition": local_RelativeHorizontalPosition, "RelativeVerticalPosition": local_RelativeVerticalPosition, "LockAnchor": local_LockAnchor, "AlternativeText": local_AlternativeText, "HasDiagram": local_HasDiagram, "HasDiagramNode": local_HasDiagramNode, "Child": local_Child, "ID": local_ID, "LayoutInCell": local_LayoutInCell, "LeftRelative": local_LeftRelative, "TopRelative": local_TopRelative, "WidthRelative": local_WidthRelative, "HeightRelative": local_HeightRelative, "RelativeHorizontalSize": local_RelativeHorizontalSize, "RelativeVerticalSize": local_RelativeVerticalSize, "ShapeStyle": local_ShapeStyle, "BackgroundStyle": local_BackgroundStyle, "Title": local_Title, "GraphicStyle": local_GraphicStyle, "Decorative": local_Decorative, }
	if (propertyName == "NoProofing"):
		retVal = this_Selection.NoProofing
		return retVal
	if (propertyName == "TopLevelTables"):
		retVal = this_Selection.TopLevelTables
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Tables", "Count": local_Count, "NestingLevel": local_NestingLevel, }
	if (propertyName == "LanguageDetected"):
		retVal = this_Selection.LanguageDetected
		return retVal
	if (propertyName == "FitTextWidth"):
		retVal = this_Selection.FitTextWidth
		return retVal
	if (propertyName == "HTMLDivisions"):
		retVal = this_Selection.HTMLDivisions
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_NestingLevel = retVal.NestingLevel
		except:
			local_NestingLevel = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "HTMLDivisions", "Count": local_Count, "NestingLevel": local_NestingLevel, }
	if (propertyName == "SmartTags"):
		retVal = this_Selection.SmartTags
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "SmartTags", "Count": local_Count, }
	if (propertyName == "ChildShapeRange"):
		retVal = this_Selection.ChildShapeRange
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		try:
			local_AutoShapeType = retVal.AutoShapeType
		except:
			local_AutoShapeType = None
		try:
			local_ConnectionSiteCount = retVal.ConnectionSiteCount
		except:
			local_ConnectionSiteCount = None
		try:
			local_Connector = retVal.Connector
		except:
			local_Connector = None
		try:
			local_Height = retVal.Height
		except:
			local_Height = None
		try:
			local_HorizontalFlip = retVal.HorizontalFlip
		except:
			local_HorizontalFlip = None
		try:
			local_Left = retVal.Left
		except:
			local_Left = None
		try:
			local_LockAspectRatio = retVal.LockAspectRatio
		except:
			local_LockAspectRatio = None
		try:
			local_Name = retVal.Name
		except:
			local_Name = None
		try:
			local_Rotation = retVal.Rotation
		except:
			local_Rotation = None
		try:
			local_Top = retVal.Top
		except:
			local_Top = None
		try:
			local_Type = retVal.Type
		except:
			local_Type = None
		try:
			local_VerticalFlip = retVal.VerticalFlip
		except:
			local_VerticalFlip = None
		try:
			local_Vertices = retVal.Vertices
		except:
			local_Vertices = None
		try:
			local_Visible = retVal.Visible
		except:
			local_Visible = None
		try:
			local_Width = retVal.Width
		except:
			local_Width = None
		try:
			local_ZOrderPosition = retVal.ZOrderPosition
		except:
			local_ZOrderPosition = None
		try:
			local_RelativeHorizontalPosition = retVal.RelativeHorizontalPosition
		except:
			local_RelativeHorizontalPosition = None
		try:
			local_RelativeVerticalPosition = retVal.RelativeVerticalPosition
		except:
			local_RelativeVerticalPosition = None
		try:
			local_LockAnchor = retVal.LockAnchor
		except:
			local_LockAnchor = None
		try:
			local_AlternativeText = retVal.AlternativeText
		except:
			local_AlternativeText = None
		try:
			local_HasDiagram = retVal.HasDiagram
		except:
			local_HasDiagram = None
		try:
			local_HasDiagramNode = retVal.HasDiagramNode
		except:
			local_HasDiagramNode = None
		try:
			local_Child = retVal.Child
		except:
			local_Child = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_LayoutInCell = retVal.LayoutInCell
		except:
			local_LayoutInCell = None
		try:
			local_LeftRelative = retVal.LeftRelative
		except:
			local_LeftRelative = None
		try:
			local_TopRelative = retVal.TopRelative
		except:
			local_TopRelative = None
		try:
			local_WidthRelative = retVal.WidthRelative
		except:
			local_WidthRelative = None
		try:
			local_HeightRelative = retVal.HeightRelative
		except:
			local_HeightRelative = None
		try:
			local_RelativeHorizontalSize = retVal.RelativeHorizontalSize
		except:
			local_RelativeHorizontalSize = None
		try:
			local_RelativeVerticalSize = retVal.RelativeVerticalSize
		except:
			local_RelativeVerticalSize = None
		try:
			local_ShapeStyle = retVal.ShapeStyle
		except:
			local_ShapeStyle = None
		try:
			local_BackgroundStyle = retVal.BackgroundStyle
		except:
			local_BackgroundStyle = None
		try:
			local_Title = retVal.Title
		except:
			local_Title = None
		try:
			local_GraphicStyle = retVal.GraphicStyle
		except:
			local_GraphicStyle = None
		try:
			local_Decorative = retVal.Decorative
		except:
			local_Decorative = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ShapeRange", "Count": local_Count, "AutoShapeType": local_AutoShapeType, "ConnectionSiteCount": local_ConnectionSiteCount, "Connector": local_Connector, "Height": local_Height, "HorizontalFlip": local_HorizontalFlip, "Left": local_Left, "LockAspectRatio": local_LockAspectRatio, "Name": local_Name, "Rotation": local_Rotation, "Top": local_Top, "Type": local_Type, "VerticalFlip": local_VerticalFlip, "Vertices": local_Vertices, "Visible": local_Visible, "Width": local_Width, "ZOrderPosition": local_ZOrderPosition, "RelativeHorizontalPosition": local_RelativeHorizontalPosition, "RelativeVerticalPosition": local_RelativeVerticalPosition, "LockAnchor": local_LockAnchor, "AlternativeText": local_AlternativeText, "HasDiagram": local_HasDiagram, "HasDiagramNode": local_HasDiagramNode, "Child": local_Child, "ID": local_ID, "LayoutInCell": local_LayoutInCell, "LeftRelative": local_LeftRelative, "TopRelative": local_TopRelative, "WidthRelative": local_WidthRelative, "HeightRelative": local_HeightRelative, "RelativeHorizontalSize": local_RelativeHorizontalSize, "RelativeVerticalSize": local_RelativeVerticalSize, "ShapeStyle": local_ShapeStyle, "BackgroundStyle": local_BackgroundStyle, "Title": local_Title, "GraphicStyle": local_GraphicStyle, "Decorative": local_Decorative, }
	if (propertyName == "HasChildShapeRange"):
		retVal = this_Selection.HasChildShapeRange
		return retVal
	if (propertyName == "FootnoteOptions"):
		retVal = this_Selection.FootnoteOptions
		try:
			local_Location = retVal.Location
		except:
			local_Location = None
		try:
			local_NumberStyle = retVal.NumberStyle
		except:
			local_NumberStyle = None
		try:
			local_StartingNumber = retVal.StartingNumber
		except:
			local_StartingNumber = None
		try:
			local_NumberingRule = retVal.NumberingRule
		except:
			local_NumberingRule = None
		try:
			local_LayoutColumns = retVal.LayoutColumns
		except:
			local_LayoutColumns = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "FootnoteOptions", "Location": local_Location, "NumberStyle": local_NumberStyle, "StartingNumber": local_StartingNumber, "NumberingRule": local_NumberingRule, "LayoutColumns": local_LayoutColumns, }
	if (propertyName == "EndnoteOptions"):
		retVal = this_Selection.EndnoteOptions
		try:
			local_Location = retVal.Location
		except:
			local_Location = None
		try:
			local_NumberStyle = retVal.NumberStyle
		except:
			local_NumberStyle = None
		try:
			local_StartingNumber = retVal.StartingNumber
		except:
			local_StartingNumber = None
		try:
			local_NumberingRule = retVal.NumberingRule
		except:
			local_NumberingRule = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "EndnoteOptions", "Location": local_Location, "NumberStyle": local_NumberStyle, "StartingNumber": local_StartingNumber, "NumberingRule": local_NumberingRule, }
	if (propertyName == "XMLNodes"):
		retVal = this_Selection.XMLNodes
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "XMLNodes", "Count": local_Count, }
	if (propertyName == "XMLParentNode"):
		retVal = this_Selection.XMLParentNode
		try:
			local_BaseName = retVal.BaseName
		except:
			local_BaseName = None
		try:
			local_Text = retVal.Text
		except:
			local_Text = None
		try:
			local_NamespaceURI = retVal.NamespaceURI
		except:
			local_NamespaceURI = None
		try:
			local_NodeType = retVal.NodeType
		except:
			local_NodeType = None
		try:
			local_NodeValue = retVal.NodeValue
		except:
			local_NodeValue = None
		try:
			local_HasChildNodes = retVal.HasChildNodes
		except:
			local_HasChildNodes = None
		try:
			local_Level = retVal.Level
		except:
			local_Level = None
		try:
			local_ValidationStatus = retVal.ValidationStatus
		except:
			local_ValidationStatus = None
		try:
			local_ValidationErrorText = retVal.ValidationErrorText
		except:
			local_ValidationErrorText = None
		try:
			local_PlaceholderText = retVal.PlaceholderText
		except:
			local_PlaceholderText = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "XMLNode", "BaseName": local_BaseName, "Text": local_Text, "NamespaceURI": local_NamespaceURI, "NodeType": local_NodeType, "NodeValue": local_NodeValue, "HasChildNodes": local_HasChildNodes, "Level": local_Level, "ValidationStatus": local_ValidationStatus, "ValidationErrorText": local_ValidationErrorText, "PlaceholderText": local_PlaceholderText, }
	if (propertyName == "Editors"):
		retVal = this_Selection.Editors
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "Editors", "Count": local_Count, }
	if (propertyName == "EnhMetaFileBits"):
		retVal = this_Selection.EnhMetaFileBits
		return retVal
	if (propertyName == "OMaths"):
		retVal = this_Selection.OMaths
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "OMaths", "Count": local_Count, }
	if (propertyName == "WordOpenXML"):
		retVal = this_Selection.WordOpenXML
		return retVal
	if (propertyName == "ContentControls"):
		retVal = this_Selection.ContentControls
		try:
			local_Count = retVal.Count
		except:
			local_Count = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ContentControls", "Count": local_Count, }
	if (propertyName == "ParentContentControl"):
		retVal = this_Selection.ParentContentControl
		try:
			local_LockContentControl = retVal.LockContentControl
		except:
			local_LockContentControl = None
		try:
			local_LockContents = retVal.LockContents
		except:
			local_LockContents = None
		try:
			local_Type = retVal.Type
		except:
			local_Type = None
		try:
			local_Title = retVal.Title
		except:
			local_Title = None
		try:
			local_DateDisplayFormat = retVal.DateDisplayFormat
		except:
			local_DateDisplayFormat = None
		try:
			local_MultiLine = retVal.MultiLine
		except:
			local_MultiLine = None
		try:
			local_Temporary = retVal.Temporary
		except:
			local_Temporary = None
		try:
			local_ID = retVal.ID
		except:
			local_ID = None
		try:
			local_ShowingPlaceholderText = retVal.ShowingPlaceholderText
		except:
			local_ShowingPlaceholderText = None
		try:
			local_DateStorageFormat = retVal.DateStorageFormat
		except:
			local_DateStorageFormat = None
		try:
			local_BuildingBlockType = retVal.BuildingBlockType
		except:
			local_BuildingBlockType = None
		try:
			local_BuildingBlockCategory = retVal.BuildingBlockCategory
		except:
			local_BuildingBlockCategory = None
		try:
			local_DateDisplayLocale = retVal.DateDisplayLocale
		except:
			local_DateDisplayLocale = None
		try:
			local_DefaultTextStyle = retVal.DefaultTextStyle
		except:
			local_DefaultTextStyle = None
		try:
			local_DateCalendarType = retVal.DateCalendarType
		except:
			local_DateCalendarType = None
		try:
			local_Tag = retVal.Tag
		except:
			local_Tag = None
		try:
			local_Checked = retVal.Checked
		except:
			local_Checked = None
		try:
			local_Color = retVal.Color
		except:
			local_Color = None
		try:
			local_Appearance = retVal.Appearance
		except:
			local_Appearance = None
		try:
			local_Level = retVal.Level
		except:
			local_Level = None
		try:
			local_RepeatingSectionItemTitle = retVal.RepeatingSectionItemTitle
		except:
			local_RepeatingSectionItemTitle = None
		try:
			local_AllowInsertDeleteSection = retVal.AllowInsertDeleteSection
		except:
			local_AllowInsertDeleteSection = None
		return { WordObjectIdKey: add_object(retVal), "__WordObjType": "ContentControl", "LockContentControl": local_LockContentControl, "LockContents": local_LockContents, "Type": local_Type, "Title": local_Title, "DateDisplayFormat": local_DateDisplayFormat, "MultiLine": local_MultiLine, "Temporary": local_Temporary, "ID": local_ID, "ShowingPlaceholderText": local_ShowingPlaceholderText, "DateStorageFormat": local_DateStorageFormat, "BuildingBlockType": local_BuildingBlockType, "BuildingBlockCategory": local_BuildingBlockCategory, "DateDisplayLocale": local_DateDisplayLocale, "DefaultTextStyle": local_DefaultTextStyle, "DateCalendarType": local_DateCalendarType, "Tag": local_Tag, "Checked": local_Checked, "Color": local_Color, "Appearance": local_Appearance, "Level": local_Level, "RepeatingSectionItemTitle": local_RepeatingSectionItemTitle, "AllowInsertDeleteSection": local_AllowInsertDeleteSection, }


# Tool: 532
@mcp.tool()
async def word_Selection_set_Property(this_Selection_wordObjId: str, propertyName: str, propertyValue):
	"""Sets properties of Selection
	
	propertyName: Name of the property. Can be one of ...
		Text, FormattedText, Start, End, Font, Style, Borders, ParagraphFormat, PageSetup, LanguageID, LanguageIDFarEast, LanguageIDOther, Flags, StartIsActive, ExtendMode, ColumnSelectMode, Orientation, NoProofing, LanguageDetected, FitTextWidth
	"""
	this_Selection = get_object(this_Selection_wordObjId)
	
	EnsureWord()
	if (propertyName == "Text"):
		this_Selection.Text = propertyValue
	if (propertyName == "FormattedText"):
		this_Selection.FormattedText = propertyValue
	if (propertyName == "Start"):
		this_Selection.Start = propertyValue
	if (propertyName == "End"):
		this_Selection.End = propertyValue
	if (propertyName == "Font"):
		this_Selection.Font = propertyValue
	if (propertyName == "Style"):
		this_Selection.Style = propertyValue
	if (propertyName == "Borders"):
		this_Selection.Borders = propertyValue
	if (propertyName == "ParagraphFormat"):
		this_Selection.ParagraphFormat = propertyValue
	if (propertyName == "PageSetup"):
		this_Selection.PageSetup = propertyValue
	if (propertyName == "LanguageID"):
		this_Selection.LanguageID = propertyValue
	if (propertyName == "LanguageIDFarEast"):
		this_Selection.LanguageIDFarEast = propertyValue
	if (propertyName == "LanguageIDOther"):
		this_Selection.LanguageIDOther = propertyValue
	if (propertyName == "Flags"):
		this_Selection.Flags = propertyValue
	if (propertyName == "StartIsActive"):
		this_Selection.StartIsActive = propertyValue
	if (propertyName == "ExtendMode"):
		this_Selection.ExtendMode = propertyValue
	if (propertyName == "ColumnSelectMode"):
		this_Selection.ColumnSelectMode = propertyValue
	if (propertyName == "Orientation"):
		this_Selection.Orientation = propertyValue
	if (propertyName == "NoProofing"):
		this_Selection.NoProofing = propertyValue
	if (propertyName == "LanguageDetected"):
		this_Selection.LanguageDetected = propertyValue
	if (propertyName == "FitTextWidth"):
		this_Selection.FitTextWidth = propertyValue




if __name__ == "__main__":
    mcp.run(transport='stdio')
