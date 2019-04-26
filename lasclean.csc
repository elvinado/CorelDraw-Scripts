REM Created On Friday, October, 05, 2018 by alvin

WITHOBJECT "CorelPHOTOPAINT.Automation.17"
	.SetDocumentInfo 1288, 9282
	.BitmapEffect "Brightness/Contrast/Intensity", CHR(7) + "BCIEffect BCIBrightness+AD0--100,BCIContrast+AD0-100,BCIIntensity+AD0-100"
	.BitmapEffect "Replace Colors", CHR(7) + "ReplaceColorsEffect RepClrInColor+AD0-5:1:1:1,RepClrOutColor+AD0-5:255:255:255,RepClrIgnoreGrayscl+AD0-0,RepClrSingleClr+AD0-0,RepClrRange+AD0-30"
	.ImageConvert 3, 1, 0, 171, 0, 45, 72, 0, 100, FALSE
END WITHOBJECT
