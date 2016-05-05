# vba-e164number

Visual Basic for Applications macro for normalizing telephone numbers to E164 format

# Install

This macro is intended for use in Microsoft Excel, and needs to be installed as a macro as follows:

- Download [E164Number.bas](E164Number.bas)
- Open your Microsoft Excel worksheet
- Open Tools -> Macro -> Visual Basic Editor, or press alt-F11.
- Choose File -> Import File, and select the `E164Number.bas` file that you downloaded (you may need to ensure the file was saved with the correct `.bas` extension).
- Close the Visual Basic Editor

# Usage

Once installed, the macro provides a function `E164Number(source, country)`. The first argument is the text that contains a telephone number, and the second is a string containing the country code (e.g. "DEU","DE" or "Germany" for Germany). If the number can be identified and correctly formated, then the result will be the number in E164 format (e.g. "+491234567890"). If it cannot, then the "#VALUE!" error will be returned.
