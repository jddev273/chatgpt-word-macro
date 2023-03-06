# ChatGPT Word Macro

A simple macro for communicating with the OpenAI ChatGPT API in Microsoft Word.  This was tested under the latest version of Microsoft Word.  Adjustments may be needed to run with older versions.

## Usage

These instructions may only apply to recent versions of Microsoft Windows.  You will need 

1. Open Microsoft Word and Click View Then Macros (or Alt-F8).
2. Enter a name (should be ChatGPT or you will need to change the macro code), make sure normal.dotm is selected in the Macros In pulldown menu then click Create.
3. The editor will open up, replace everything with the code in the macro.bas file.
4. Replace `sk-YOUR-CHATGPT-KEY-HERE` with your own OpenAI API key.
5. Run the `ChatGPT` subroutine by pressing Alt-F8 and clicking run.
6. Enter the prompt you want to send to the ChatGPT API in the input box that appears.
7. The response from the ChatGPT API will be inserted into the document where the cursor was last placed.

## Troubleshooting

If you receive an error related to `WinHttp.WinHttpRequest.5.1`, it means that the `WinHttp` library is not registered on your computer. To resolve this issue, you can try re-registering the library by running the following command in an elevated command prompt:

regsvr32 %systemroot%\system32\winhttp.dll

If the error persists, you may need to reinstall the library. You can download it from the Microsoft website.

## Note
The UnescapeString function is used to format the response from the API into a human-readable format. If you encounter any errors or unexpected output, please refer to the comments in the code for further assistance.

## Disclaimer
The authors and contributors of this program provide it as-is, without any warranties or guarantees. They cannot be held responsible for any damages resulting from the use of this program.

## License
This program is licensed under the MIT license.

## Author
Johann Dowa