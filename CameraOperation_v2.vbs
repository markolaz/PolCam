Function Filt(Color)

	Select Case LCase(Color)

		case "red", "r"
			Filt = 0

		case "green", "g"
			Filt = 1

		case "blue", "b"
			Filt = 2

		case "none", "n"
			Filt = 3

		case "", "quit", "q"
			Exit Function

		case Else
			'WScript.Echo "Not an applicable filter option, please try again."
			Color = InputBox("Inapplicable Filter Color Entered, Try Again. Enter Filter Color (red/green/blue/none)")
			Filt = Filt(Color)

	End Select

End Function

Function Shutter(Shutt)

	Select case LCase(Shutt)
	
		case "yes", "y"
			Shutter = 1
	
		case "no", "n"
			Shutter = 0

		case "", "quit", "q"
			Exit Function
	
		case Else
			'WScript.Echo "Not an applicable shutter option, please try again."
			Shutt = InputBox("Inapplicable Shutter Option Entered, Try Again. vbCrLf Open Shutter? (y/n)")
			Shutter = Shutter(Shutt)
	
	End Select

End Function

Sub CamExpose()

Target = InputBox("Enter Name of Target or Target Identifier")

Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FolderExists("C:\Users\Marko\Desktop\Polarization Camera\Data\" & Target) Then
Else
    objFSO.CreateFolder("C:\Users\Marko\Desktop\Polarization Camera\Data\" & Target)
End If

Prefix = InputBox("Enter Filename Prefix for Dataset")

Shutt = InputBox("Open Shutter? (y/n)")

Shuttering = Shutter(Shutt)

ExpTime = InputBox("Enter Exposure Time in Seconds")

Color = InputBox("Enter Filter Color (red/green/blue/none)")

Filtering = Filt(Color)

NImages = InputBox("Enter Integer Number of Images to be Taken")

Set Camera = CreateObject("MaxIm.CCDCamera")

Set Autofocuser = CreateObject("MaxIm.Application")

Camera.LinkEnabled = True
Autofocuser.FocuserConnected = True

if Not Camera.LinkEnabled Then
	WScript.Echo "Failed to connect to camera!"
	Quit
End If

if Not Autofocuser.FocuserConnected Then
   WScript.Echo "Failed to connect to focuser!"
   'Quit
End If

Autofocuser.AutoFocus()

Do While Not AutoFocusStatus
WScript.Sleep 1000
Loop

For Image = 1 to NImages

	WScript.Echo "Camera is ready, exposing..."

	Camera.Expose ExpTime, Shuttering, Filtering

	Do While Not Camera.ImageReady
	Loop

	Camera.SaveImage "C:\Users\Marko\Desktop\Polarization Camera\Data\" & Target & "\" & Prefix & Image & ".fits"

	WScript.Echo "Exposure finished, saved image as " & Target & "\" & Prefix & Image & ".fits out of " & NImages & " images"

Next

End Sub

CamExpose()



' make subroutines for each of the input statements
' make it such that "enter" uses the default input, place global variables at top that can be set by user beforehand for repeated exposures
' make checks for exposure time (negative), filename (length, etc.), and the rest
' make it so that existing files cannot be overwritten



'For Image = 1 To NImages
''	WScript.Echo Image, ExpTime, Filt
'Next

'Select case LCase(Filt)
'
'	case "red"
'		Filt = 0
'
'	case "green"
'		Filt = 1 
'
'	case "blue"
'		Filt = 2
'
'	case Else
'		WScript.Echo "Not an applicable filter option, please try again."
'
'End select

'Select case LCase(Shutter)
'
'	case "yes", "y"
'		Shutter = 1
'
'	case "no", "n"
'		Shutter = 0
'
'	case Else
'		WScript.Echo "Not an applicable shutter option, please try again."
'
'End Select

'WScript.Echo "This Image Number is" & Image

'Wscript.Echo ExpTime
'Wscript.Echo Filt

'Wscript.Echo "Hello"
'Wscript.Quit 0