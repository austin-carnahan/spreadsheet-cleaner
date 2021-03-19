How do I install the custom geocoding functions so they're available in all my workbooks?

!!! Please see the installation video for a full walkthough of the steps outlined below !!!

1. Open a new Excel workbook

2. Press Alt+F11

3. File -> Import File -> then import the following three files provided by USDR:
	1. CustomGeocodingFunctions.bas
	2. Dictionary.cls
	3. JsonConverter.bas

4. Tools -> References -> then check the box for "Microsoft XML, v6.0" -> OK

5. Save the workbook as an Excel Add-in here: 

	!!! Remember to replace the "[USER]" in the path below with your real username !!!

	C:\Users\[USER]\AppData\Roaming\Microsoft\Excel\XLSTART

6. File -> Options -> Add-ins -> Manage [Excel Add-ins] -> check the box next to the custom add-in.


You're done! You should be able to use the custom geocoding functions in any workbook you open!
