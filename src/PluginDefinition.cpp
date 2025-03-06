//this file is part of notepad++
//Copyright (C)2022 Don HO <don.h@free.fr>
//
//This program is free software; you can redistribute it and/or
//modify it under the terms of the GNU General Public License
//as published by the Free Software Foundation; either
//version 2 of the License, or (at your option) any later version.
//
//This program is distributed in the hope that it will be useful,
//but WITHOUT ANY WARRANTY; without even the implied warranty of
//MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//GNU General Public License for more details.
//
//You should have received a copy of the GNU General Public License
//along with this program; if not, write to the Free Software
//Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.

#include "PluginDefinition.h"
#include "DockingFeature/LoaderDlg.h"
#include "DockingFeature/ChatSettingsDlg.h"
#include "menuCmdID.h"

// For file + cURL + JSON ops
#include <wchar.h>
#include <shlwapi.h>
#include <curl/curl.h>
#include <codecvt> // codecvt_utf8
#include <locale>  // wstring_convert
#include <nlohmann/json.hpp>
#include <regex>

// For "async" cURL calls
#include <thread>

// Instead of `#include <commctrl.h>` we define the required constants only!
#define UD_MAXVAL 0x7fff // 32767 (more than enough)

// For cURL JSON requests/responses
using json = nlohmann::json;

// Loader window ("Please wait for Ollama's response…")
HANDLE _hModule;
LoaderDlg _loaderDlg;
ChatSettingsDlg _chatSettingsDlg;

// Config file related vars/constants
TCHAR iniFilePath[MAX_PATH];
TCHAR instructionsFilePath[MAX_PATH]; // Aka. file for Ollama system message

// The plugin data that Notepad++ needs
FuncItem funcItem[nbFunc];

// The data of Notepad++ that you can use in your plugin commands
NppData nppData;

// Config file related vars
std::wstring configAPIValue_secretKey        = TEXT("not-required-for-ollama"); // No API key needed for local Ollama
std::wstring configAPIValue_baseURL          = TEXT("http://localhost:11434/"); // Default Ollama API endpoint
std::wstring configAPIValue_proxyURL         = TEXT("0"); // 0: don't use proxy. Trailing '/' will be erased (if any)
std::wstring configAPIValue_model            = TEXT("llama3"); // Default Ollama model
std::wstring configAPIValue_instructions     = TEXT(""); // System message ("instuctions") for the Ollama API e.g. "Translate the given text into English." or "Create a PHP function based on the received text.". Leave empty to skip.
std::wstring configAPIValue_temperature      = TEXT("0.7");
std::wstring configAPIValue_maxTokens        = TEXT("0"); // 0: Skip `max_tokens` API setting. Recommended max. value:  <4.000
std::wstring configAPIValue_topP             = TEXT("0.8");
std::wstring configAPIValue_frequencyPenalty = TEXT("0");
std::wstring configAPIValue_presencePenalty  = TEXT("0");
bool isKeepQuestion                          = true;
std::vector<std::wstring> chatHistory        = {};

// Collect selected text by Scintilla here
TCHAR selectedText[9999];

//
// Initialize your plugin data here
// It will be called while plugin loading
void pluginInit(HANDLE hModule)
{
	_hModule = hModule;
	_loaderDlg.init((HINSTANCE)_hModule, nppData._nppHandle);

	// Init Chat Settings modal dialog
	_chatSettingsDlg.init((HINSTANCE)_hModule, nppData._nppHandle);
	_chatSettingsDlg.chatSetting_isChat = false;
	_chatSettingsDlg.chatSetting_chatLimit = 10;
}

//
// Here you can do the clean up, save the parameters (if any) for the next session
//
void pluginCleanUp()
{
	wchar_t chatLimitBuffer[6];
	wsprintfW(chatLimitBuffer, L"%d", _chatSettingsDlg.chatSetting_chatLimit);
	::WritePrivateProfileString(TEXT("PLUGIN"), TEXT("keep_question"), isKeepQuestion ? TEXT("1") : TEXT("0"), iniFilePath);
	::WritePrivateProfileString(TEXT("PLUGIN"), TEXT("is_chat"), _chatSettingsDlg.chatSetting_isChat ? TEXT("1") : TEXT("0"), iniFilePath);
	::WritePrivateProfileString(TEXT("PLUGIN"), TEXT("chat_limit"), chatLimitBuffer, iniFilePath); // Convert int (9+) to LPCWSTR
}

//
// Initialization of your plugin commands
// You should fill your plugins commands here
void commandMenuInit()
{
	TCHAR configDirPath[MAX_PATH];

	// Get path to the plugin config + instructions (aka. Ollama system message) file
	::SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH, (LPARAM)configDirPath);

	// If config path doesn't exist, create it
	if (PathFileExistsW(configDirPath) == FALSE) // Modified from `PathFileExists()`
	{
		::CreateDirectory(configDirPath, NULL);
	}

	// Prepare config + instructions (aka. system message) file
	PathCombine(iniFilePath, configDirPath, TEXT("NppOpenAI.ini"));
	PathCombine(instructionsFilePath, configDirPath, TEXT("NppOpenAI_instructions"));

	// Load config file content
	loadConfig(true);


    //--------------------------------------------//
    //-- STEP 3. CUSTOMIZE YOUR PLUGIN COMMANDS --//
    //--------------------------------------------//
    // with function :
    // setCommand(int index,                      // zero based number to indicate the order of command
    //            TCHAR *commandName,             // the command name that you want to see in plugin menu
    //            PFUNCPLUGINCMD functionPointer, // the symbol of function (function pointer) associated with this command. The body should be defined below. See Step 4.
    //            ShortcutKey *shortcut,          // optional. Define a shortcut to trigger this command
    //            bool check0nInit                // optional. Make this menu item be checked visually
    //            );

	// Shortcuts for Ollama plugin
	ShortcutKey *askOllamaKey = new ShortcutKey;
	askOllamaKey->_isAlt = false;
	askOllamaKey->_isCtrl = true;
	askOllamaKey->_isShift = true;
	askOllamaKey->_key = 0x4f; // 'O'

	// Plugin menu items
	setCommand(0, TEXT("Ask &Ollama"), askChatGPT, askOllamaKey, false);
	setCommand(1, TEXT("---"), NULL, NULL, false);
	setCommand(2, TEXT("&Edit Config"), openConfig, NULL, false);
	setCommand(3, TEXT("Edit &Instructions"), openInsturctions, NULL, false);
	setCommand(4, TEXT("&Load Config"), loadConfigWithoutPluginSettings, NULL, false);
	setCommand(5, TEXT("---"), NULL, NULL, false);
	setCommand(6, TEXT("&Keep my question"), keepQuestionToggler, NULL, isKeepQuestion);
	setCommand(7, TEXT("NppOllama &Chat Settings"), openChatSettingsDlg, NULL, false); // Text will be updated by `updateToolbarIcons()` » `updateChatSettings()`
	setCommand(8, TEXT("---"), NULL, NULL, false);
	setCommand(9, TEXT("&About"), openAboutDlg, NULL, false);
}

// Add/update toolbar icons
void updateToolbarIcons()
{

	// Prepare icons to open Chat Settings
	int hToolbarBmp = IDB_PLUGINNPPOPENAI_TOOLBAR_CHAT;
	int hToolbarIcon = IDI_PLUGINNPPOPENAI_TOOLBAR_CHAT;
	int hToolbarIconDarkMode = IDI_PLUGINNPPOPENAI_TOOLBAR_CHAT_DM;
	/* // TODO: update toolbar icons on-the-fly (turning chat on/off or reaching chat limit)
	if (!_chatSettingsDlg.chatSetting_isChat || _chatSettingsDlg.chatSetting_chatLimit == 0)
	{
		hToolbarBmp = IDB_PLUGINNPPOPENAI_TOOLBAR_NO_CHAT;
		hToolbarIcon = IDI_PLUGINNPPOPENAI_TOOLBAR_NO_CHAT;
		hToolbarIconDarkMode = IDI_PLUGINNPPOPENAI_TOOLBAR_NO_CHAT_DM;
	}
	// */

	// Send Chat Settings icons to Notepad++
	toolbarIconsWithDarkMode chatSettingsIcons;
	chatSettingsIcons.hToolbarBmp = ::LoadBitmap((HINSTANCE)_hModule, MAKEINTRESOURCE(hToolbarBmp));
	chatSettingsIcons.hToolbarIcon = ::LoadIcon((HINSTANCE)_hModule, MAKEINTRESOURCE(hToolbarIcon));
	chatSettingsIcons.hToolbarIconDarkMode = ::LoadIcon((HINSTANCE)_hModule, MAKEINTRESOURCE(hToolbarIconDarkMode));
	::SendMessage(nppData._nppHandle, NPPM_ADDTOOLBARICON_FORDARKMODE, funcItem[7]._cmdID, (LPARAM)&chatSettingsIcons); // Open Chat Settings
	updateChatSettings();
}

//
// Here you can do the clean up (especially for the shortcut)
//
void commandMenuCleanUp()
{
	// Don't forget to deallocate your shortcut here
	delete funcItem[0]._pShKey;
	_loaderDlg.destroy();
	_chatSettingsDlg.destroy();
}


//
// This function help you to initialize your plugin commands
//
bool setCommand(size_t index, TCHAR *cmdName, PFUNCPLUGINCMD pFunc, ShortcutKey *sk, bool check0nInit) 
{
    if (index >= nbFunc)
        return false;

    if (!pFunc)
        return false;

    lstrcpy(funcItem[index]._itemName, cmdName);
    funcItem[index]._pFunc = pFunc;
    funcItem[index]._init2Check = check0nInit;
    funcItem[index]._pShKey = sk;

    return true;
}


//----------------------------------------------//
//-- STEP 4. DEFINE YOUR ASSOCIATED FUNCTIONS --//
//----------------------------------------------//

// Wrapper function to match the PFUNCPLUGINCMD signature
void loadConfigWithoutPluginSettings() {
	loadConfig(false);
}

// Load (and create if not found) config file
void loadConfig(bool loadPluginSettings)
{
	wchar_t tbuffer2[256];
	FILE* instructionsFile;

	// Set up a default plugin config (if necessary)
	if (::GetPrivateProfileString(TEXT("API"), TEXT("model"), NULL, tbuffer2, 128, iniFilePath) == NULL)
	{
		::WritePrivateProfileString(TEXT("INFO"), TEXT("; === USING OLLAMA ON LOCALHOST =="), TEXT(""), iniFilePath);
		::WritePrivateProfileString(TEXT("INFO"), TEXT("; == Default model is llama3, but you can use any model you have pulled in Ollama ="), TEXT(""), iniFilePath);
		::WritePrivateProfileString(TEXT("INFO"), TEXT("; == You don't need an API key for Ollama running on localhost ="), TEXT(""), iniFilePath);
		::WritePrivateProfileString(TEXT("INFO"), TEXT("; == Token and pricing info: Ollama is free and runs locally ="), TEXT(""), iniFilePath);
		::WritePrivateProfileString(TEXT("INFO"), TEXT("; == Set `max_tokens=0` to skip this setting. ="), TEXT(""), iniFilePath);
		::WritePrivateProfileString(TEXT("API"), TEXT("secret_key"), configAPIValue_secretKey.c_str(), iniFilePath);
		::WritePrivateProfileString(TEXT("API"), TEXT("model"), configAPIValue_model.c_str(), iniFilePath);
		::WritePrivateProfileString(TEXT("API"), TEXT("temperature"), configAPIValue_temperature.c_str(), iniFilePath);
		::WritePrivateProfileString(TEXT("API"), TEXT("max_tokens"), configAPIValue_maxTokens.c_str(), iniFilePath);
		::WritePrivateProfileString(TEXT("API"), TEXT("top_p"), configAPIValue_topP.c_str(), iniFilePath);
		::WritePrivateProfileString(TEXT("API"), TEXT("frequency_penalty"), configAPIValue_frequencyPenalty.c_str(), iniFilePath);
		::WritePrivateProfileString(TEXT("API"), TEXT("presence_penalty"), configAPIValue_presencePenalty.c_str(), iniFilePath);
		::WritePrivateProfileString(TEXT("PLUGIN"), TEXT("keep_question"), TEXT("1"), iniFilePath);
		::WritePrivateProfileString(TEXT("PLUGIN"), TEXT("total_tokens_used"), TEXT("0"), iniFilePath);
	}

	// Set up the default API URL
	if (::GetPrivateProfileString(TEXT("API"), TEXT("api_url"), NULL, tbuffer2, 256, iniFilePath) == NULL)
	{
		::WritePrivateProfileString(TEXT("API"), TEXT("api_url"), configAPIValue_baseURL.c_str(), iniFilePath);
		::WritePrivateProfileString(TEXT("INFO"), TEXT("; == The endpoint for Ollama is /api/generate. If you're running Ollama on a different port, change the URL. ="), TEXT(""), iniFilePath);
	}

	// Chat preparations + create file for instructions (aka. system message)
	if (::GetPrivateProfileString(TEXT("PLUGIN"), TEXT("is_chat"), NULL, tbuffer2, 2, iniFilePath) == NULL)
	{
		::WritePrivateProfileString(TEXT("PLUGIN"), TEXT("is_chat"), TEXT("0"), iniFilePath);
		if ((instructionsFile = _wfopen(instructionsFilePath, L"w, ccs=UNICODE")) != NULL)
		{
			fclose(instructionsFile);
		}
		else
		{
			instructionsFileError(TEXT("The instructions (system message) file could not be created:\n\n"), TEXT("NppOllama: unavailable instructions file"));
		}
	}

	// Set up the default chat settings
	if (::GetPrivateProfileString(TEXT("PLUGIN"), TEXT("chat_limit"), NULL, tbuffer2, 2, iniFilePath) == NULL)
	{
		::WritePrivateProfileString(TEXT("PLUGIN"), TEXT("chat_limit"), TEXT("10"), iniFilePath);
		::WritePrivateProfileString(TEXT("INFO"), TEXT("; == Use `instructions` to set a system message for the Ollama API e.g. 'Translate the given message into English.' or 'Create a PHP function based on the received text.' Optional, leave empty to skip. ="), TEXT(""), iniFilePath);
	}

	// Set up proxy settings
	if (::GetPrivateProfileString(TEXT("API"), TEXT("proxy_url"), NULL, tbuffer2, 256, iniFilePath) == NULL)
	{
		::WritePrivateProfileString(TEXT("API"), TEXT("proxy_url"), configAPIValue_proxyURL.c_str(), iniFilePath);
		::WritePrivateProfileString(TEXT("INFO"), TEXT("; == Enter a `proxy_url` to use proxy like 'http://127.0.0.1:80'. Optional, enter 0 (zero) to skip. ="), TEXT(""), iniFilePath);
	}

	// Get instructions (aka. system message) file
	if ((instructionsFile = _wfopen(instructionsFilePath, L"r, ccs=UNICODE")) != NULL)
	{
		wchar_t instructionsBuffer[9999];
		configAPIValue_instructions = TEXT("");
		while (fgetws(instructionsBuffer, 9999, instructionsFile))
		{
			configAPIValue_instructions += instructionsBuffer;
		}
		fclose(instructionsFile);
	}
	else
	{
		instructionsFileError(TEXT("The instructions (system message) file was not found:\n\n"), TEXT("NppOllama: missing instructions file"));
	}

	// Get API config/settings
	::GetPrivateProfileString(TEXT("API"), TEXT("secret_key"), NULL, tbuffer2, 256, iniFilePath);
	configAPIValue_secretKey = std::wstring(tbuffer2);

	::GetPrivateProfileString(TEXT("API"), TEXT("api_url"), NULL, tbuffer2, 256, iniFilePath);
	configAPIValue_baseURL = std::wstring(tbuffer2);

	::GetPrivateProfileString(TEXT("API"), TEXT("proxy_url"), NULL, tbuffer2, 256, iniFilePath);
	configAPIValue_proxyURL = std::wstring(tbuffer2);

	::GetPrivateProfileString(TEXT("API"), TEXT("model"), NULL, tbuffer2, 128, iniFilePath);
	if (std::wstring(tbuffer2) == TEXT("gpt-4"))
	{
		::WritePrivateProfileString(TEXT("API"), TEXT("model"), configAPIValue_model.c_str(), iniFilePath);
		::GetPrivateProfileString(TEXT("API"), TEXT("model"), NULL, tbuffer2, 128, iniFilePath);
	}
	configAPIValue_model = std::wstring(tbuffer2);

	::GetPrivateProfileString(TEXT("API"), TEXT("temperature"), NULL, tbuffer2, 16, iniFilePath);
	configAPIValue_temperature = std::wstring(tbuffer2);

	::GetPrivateProfileString(TEXT("API"), TEXT("max_tokens"), NULL, tbuffer2, 5, iniFilePath);
	configAPIValue_maxTokens = std::wstring(tbuffer2);

	::GetPrivateProfileString(TEXT("API"), TEXT("top_p"), NULL, tbuffer2, 5, iniFilePath);
	configAPIValue_topP = std::wstring(tbuffer2);

	::GetPrivateProfileString(TEXT("API"), TEXT("frequency_penalty"), NULL, tbuffer2, 5, iniFilePath);
	configAPIValue_frequencyPenalty = std::wstring(tbuffer2);

	::GetPrivateProfileString(TEXT("API"), TEXT("presence_penalty"), NULL, tbuffer2, 5, iniFilePath);
	configAPIValue_presencePenalty = std::wstring(tbuffer2);

	// Get Plugin config/settings
	// Do NOT load "PLUGIN" section when clicking the Load Config menu item (may cause misconfiguration)
	if (loadPluginSettings)
	{
		isKeepQuestion = (::GetPrivateProfileInt(TEXT("PLUGIN"), TEXT("keep_question"), 1, iniFilePath) != 0);
		_chatSettingsDlg.chatSetting_isChat = (::GetPrivateProfileInt(TEXT("PLUGIN"), TEXT("is_chat"), 0, iniFilePath) != 0);
		int tmpChatLimit = ::GetPrivateProfileInt(TEXT("PLUGIN"), TEXT("chat_limit"), _chatSettingsDlg.chatSetting_chatLimit, iniFilePath);
		_chatSettingsDlg.chatSetting_chatLimit = (tmpChatLimit <= 0)
			? 1 // Chat limit: min. value
			: ((tmpChatLimit > UD_MAXVAL)
				? UD_MAXVAL // Chat limit: max. value
				: tmpChatLimit);

		// Update chat menu item text (if already initialized)
		if (funcItem[7]._pFunc)
		{
			updateChatSettings();
		}
	}
}

// Call Ollama API
void askChatGPT()
{
	// Get current Scintilla
	long currentEdit;
	::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&currentEdit);
	HWND curScintilla = (currentEdit == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;

	// Get current selection
	size_t selstart  = ::SendMessage(curScintilla, SCI_GETSELECTIONSTART, 0, 0);
	size_t selend    = ::SendMessage(curScintilla, SCI_GETSELECTIONEND, 0, 0);
	size_t sellength = selend - selstart;

	// Check if everything is fine
	bool isEditable  = !(int)::SendMessage(curScintilla, SCI_GETREADONLY, 0, 0);
	if (isEditable
		&& selend > selstart
		&& sellength < 2048
		&& ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTWORD, 2048, (LPARAM)selectedText)
		)
	{
		// Data to post via cURL - Ollama format
		json postData = {
			{"model", toUTF8(configAPIValue_model)},
			{"temperature", std::stod(configAPIValue_temperature)},
			{"top_p", std::stod(configAPIValue_topP)}
		};

		// Add `max_tokens` setting if specified
		if (std::stoi(configAPIValue_maxTokens) > 0)
		{
			postData[toUTF8(TEXT("max_tokens"))] = std::stoi(configAPIValue_maxTokens);
		}

		// Add system instructions if available
		if (!configAPIValue_instructions.empty())
		{
			postData["system"] = toUTF8(configAPIValue_instructions);
		}

		// Add the main prompt
		postData["prompt"] = toUTF8(selectedText);
		
		// Update URLs for API call
		bool isReady2CallOllama = true;
		std::string OpenAIURL = toUTF8(configAPIValue_baseURL).erase(toUTF8(configAPIValue_baseURL).find_last_not_of("/") + 1);
		std::string ProxyURL = toUTF8(configAPIValue_proxyURL).erase(toUTF8(configAPIValue_proxyURL).find_last_not_of("/") + 1);
		
		// Set the Ollama API endpoint
		OpenAIURL += "/api/generate";

		// Ready to call Ollama
		if (isReady2CallOllama)
		{
			// Create/Show a loader dialog ("Please wait..."), disable main window
			_loaderDlg.doDialog();
			::EnableWindow(nppData._nppHandle, FALSE);

			// Prepare to start a new thread
			auto curlLambda = [](std::string OpenAIURL, std::string ProxyURL, json postData, HWND curScintilla)
			{
				std::string JSONRequest = postData.dump();

				// Try to call Ollama and store the results in `JSONBuffer`
				std::string JSONBuffer;
				bool isSuccessCall = callOpenAI(OpenAIURL, ProxyURL, JSONRequest, JSONBuffer);

				// Hide loader dialog, enable main window
				_loaderDlg.display(false);
				::EnableWindow(nppData._nppHandle, TRUE);
				::SetForegroundWindow(nppData._nppHandle);

				// Return if something went wrong
				if (!isSuccessCall)
				{
					return;
				}

				// Parse response
				try
				{
					json JSONResponse = json::parse(JSONBuffer);

					// Handle Ollama response format
					if (JSONResponse.contains("response"))
					{
						// Get the response text
						std::string responseText;
						JSONResponse["response"].get_to(responseText);

						// Replace selected text with response in the main Notepad++ window
						replaceSelected(curScintilla, responseText);

						// Update chat history
						chatHistory.push_back(selectedText);
						chatHistory.push_back(std::wstring(responseText.begin(), responseText.end()));

						// No need to update token counts for Ollama as it doesn't track them
					}
					else if (JSONResponse.contains("error"))
					{
						std::string errorResponse;
						TCHAR errorResponseWide[512] = { 0, };
						JSONResponse["error"].get_to(errorResponse);
						std::copy(errorResponse.begin(), errorResponse.end(), errorResponseWide);
						::MessageBox(nppData._nppHandle, errorResponseWide, TEXT("Ollama: Error response"), MB_ICONEXCLAMATION);
					}
					else
					{
						::MessageBox(nppData._nppHandle, TEXT("Missing 'response' in JSON response!"), TEXT("Ollama: Invalid answer"), MB_ICONEXCLAMATION);
					}
				}
				catch (json::parse_error& ex)
				{
					std::string responseException = ex.what();
					std::string responseText = JSONBuffer.c_str();
					responseText += "\n\n" + responseException;
					replaceSelected(curScintilla, responseText);
					::MessageBox(nppData._nppHandle, TEXT("Invalid or non-JSON response!\n\nSee details in the main window"), TEXT("Ollama: Invalid response"), MB_ICONERROR);
				}
			};

			std::thread curlThread(curlLambda, OpenAIURL, ProxyURL, postData, curScintilla);
			curlThread.detach();
		}
	}
	else if (!isEditable)
	{
		::MessageBox(nppData._nppHandle, TEXT("This file is not editable"), TEXT("Ollama: Invalid file"), MB_ICONERROR);
	}
	else if (selend <= selstart)
	{
		::MessageBox(nppData._nppHandle, TEXT("Please select a text first"), TEXT("Ollama: Missing question"), MB_ICONWARNING);
	}
	else if (sellength >= 2048)
	{
		::MessageBox(nppData._nppHandle, TEXT("The selected text is too long. Please select a text shorter than 2048 characters"), TEXT("Ollama: Invalid question"), MB_ICONWARNING);
	}
	else
	{
		::MessageBox(nppData._nppHandle, TEXT("Please try to select a question first"), TEXT("Ollama: Unknown error"), MB_ICONERROR);
	}
}

// Call Ollama via cURL
bool callOpenAI(std::string OpenAIURL, std::string ProxyURL, std::string JSONRequest, std::string& JSONResponse)
{
	// Prepare cURL
	CURL* curl;
	CURLcode res;
	curl_global_init(CURL_GLOBAL_ALL); // In windows, this will init the winsock stuff

	// Get a cURL handle
	curl = curl_easy_init();
	if (!curl)
	{
		return false;
	}

	// Get the CA bundle file for cURL
	TCHAR CACertFilePath[MAX_PATH];
	const TCHAR CACertFileName[] = TEXT("NppOpenAI\\cacert.pem");
	::SendMessage(nppData._nppHandle, NPPM_GETPLUGINHOMEPATH, MAX_PATH, (LPARAM)CACertFilePath);
	PathAppend(CACertFilePath, CACertFileName);

	// Prepare cURL SetOpts
	struct curl_slist* headerList = NULL;
	headerList = curl_slist_append(headerList, "Content-Type: application/json");
	char userAgent[255];
	sprintf(userAgent, "NppOllama/%s", NPPOPENAI_VERSION);

	// cURL SetOpts
	curl_easy_setopt(curl, CURLOPT_URL, OpenAIURL.c_str());
	if (ProxyURL != "" && ProxyURL != "0")
	{
		curl_easy_setopt(curl, CURLOPT_PROXY, ProxyURL.c_str());
	}
	curl_easy_setopt(curl, CURLOPT_FOLLOWLOCATION, 1L);
	curl_easy_setopt(curl, CURLOPT_HTTPPROXYTUNNEL, 1L); // Corp. proxies etc.
	curl_easy_setopt(curl, CURLOPT_CAINFO, toUTF8(CACertFilePath).c_str());
	curl_easy_setopt(curl, CURLOPT_USERAGENT, userAgent);
	curl_easy_setopt(curl, CURLOPT_HTTPHEADER, headerList);
	curl_easy_setopt(curl, CURLOPT_POSTFIELDS, JSONRequest.c_str());
	curl_easy_setopt(curl, CURLOPT_POST, 1L);
	curl_easy_setopt(curl, CURLOPT_WRITEDATA, &JSONResponse);
	curl_easy_setopt(curl, CURLOPT_WRITEFUNCTION, OpenAIcURLCallback); // Send all data to this function

	// Perform the request, res will get the return code
	res = curl_easy_perform(curl);
	bool isCurlOK = (res == CURLE_OK);

	// Handle response + check for errors
	if (!isCurlOK)
	{
		char curl_error[512];
		sprintf(curl_error, "An error occurred while accessing the Ollama server:\n%s", curl_easy_strerror(res));
		::MessageBox(nppData._nppHandle, myMultiByteToWideChar(curl_error), TEXT("Ollama: Connection Error"), MB_ICONERROR);
	}

	// Cleanup (including headers)
	curl_easy_cleanup(curl);
	curl_slist_free_all(headerList);
	return isCurlOK;
}

// Open config file
void openConfig()
{
	::SendMessage(nppData._nppHandle, NPPM_DOOPEN, 0, (LPARAM)iniFilePath);
}

// Open insturctions file
void openInsturctions()
{
	::SendMessage(nppData._nppHandle, NPPM_DOOPEN, 0, (LPARAM)instructionsFilePath);
	::MessageBox(nppData._nppHandle, TEXT("You can give instructions (system message) to Ollama here, e.g.: Translate the received text into English.\n\n\
Leave the file empty to skip the instructions.\n\n\
After saving this file, don't forget to click Plugins » NppOllama » Load Config to apply the changes."), TEXT("NppOllama: Instructions"), MB_ICONINFORMATION);
}

// Toggle "Keep my question" menu item
void keepQuestionToggler()
{
	isKeepQuestion = !isKeepQuestion;
	::CheckMenuItem(::GetMenu(nppData._nppHandle), funcItem[6]._cmdID, MF_BYCOMMAND | (isKeepQuestion ? MF_CHECKED : MF_UNCHECKED));
}

// Open Chat Settings dialog
void openChatSettingsDlg()
{
	_chatSettingsDlg.doDialog();
}

// Update chat settings UI
void updateChatSettings(bool isWriteToFile)
{
	HMENU chatMenu = ::GetMenu(nppData._nppHandle);
