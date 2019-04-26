REM Created On Wednesday, August, 15, 2018 by alvin

WITHOBJECT "CorelPHOTOPAINT.Automation.17"
	.SetDocumentInfo 2246, 3177
	.BitmapEffect "Brightness/Contrast/Intensity", CHR(7) + "BCIEffect BCIBrightness+AD0--100,BCIContrast+AD0-80,BCIIntensity+AD0-90"
	.ImageConvert 3, 1, 0, 235, 0, 45, 72, 0, 100, FALSE
END WITHOBJECT
