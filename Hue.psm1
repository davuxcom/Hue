# Notes:
#      - Serial requests are faster than parallel, the bridge has little power.
function Get-HueLight($BridgeIP = (gs Hue.BridgeIP), 
                      $UserName = (gs Hue.UserName),
                      [switch]$Groups,
                      [switch]$All) {

    $Service = $(if ($Groups -or $All) { "groups" } else { "lights" })

	$Lights = Invoke-RestMethod "http://$BridgeIP/api/$UserName/$Service" | 
        Get-Member | ? { $_.MemberType -eq "NoteProperty" } |  select -exp Name | % {
                [PSCustomObject]@{'ID'=[int]::Parse($_)
                                  'Name'=$pack.$_.Name}
            } | sort ID

    if ($Groups -or $All) { # Add the all-groups group (as of API 1 this is the only one anyway)
        $Lights += @([PSCustomObject]@{'ID'=0
                                       'Name'="All Lights"})
    }

    $Lights | % {
        $Light = Invoke-RestMethod "http://$BridgeIP/api/$UserName/$Service/$($_.ID)"
        [PsCustomObject]@{
            'ID'=$_.ID
            'Name'=$Light.name
            'On'=$Light.state.on
            'Brightness'=$Light.state.bri
            'Hue'=$Light.state.hue
            'Effect'=$Light.state.effect
            'Alert'=$Light.state.alert
            'Saturation'=$Light.state.sat
            'Reachable'=$Light.state.reachable
            'xy'=$Light.state.xy # TODO: implement Set- for xy/ct
            'ct'=$Light.state.ct
            'IsGroup'=($Groups -or $All)
        }
    }
}

function Set-HueLight([Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
                      $Light,
                      $BridgeIP = (gs Hue.BridgeIP), 
                      $UserName = (gs Hue.userName),
                      [switch]$On,
                      [switch]$Off,
                      [switch]$AsScheduleCommand,
                      [switch]$Relax,
                      [switch]$White,
                      [switch]$Default,
                      [int]$Brightness = -1,
                      [int]$Hue = -1,
                      [int]$Saturation = -1,
                      [int]$TransitionTime = -1,
                      [ValidateSet("AliceBlue", "AntiqueWhite", "Aqua", "Aquamarine", "Azure", "Beige", "Bisque", "Black", "BlanchedAlmond", "Blue", "BlueViolet", "Brown", "BurlyWood", "CadetBlue", "Chartreuse", "Chocolate", "Coral", "CornflowerBlue", "Cornsilk", "Crimson", "Cyan", "DarkBlue", "DarkCyan", "DarkGoldenrod", "DarkGray", "DarkGreen", "DarkKhaki", "DarkMagenta", "DarkOliveGreen", "DarkOrange", "DarkOrchid", "DarkRed", "DarkSalmon", "DarkSeaGreen", "DarkSlateBlue", "DarkSlateGray", "DarkTurquoise", "DarkViolet", "DeepPink", "DeepSkyBlue", "DimGray", "DodgerBlue", "Firebrick", "FloralWhite", "ForestGreen", "Fuchsia", "Gainsboro", "GhostWhite", "Gold", "Goldenrod", "Gray", "Green", "GreenYellow", "Honeydew", "HotPink", "IndianRed", "Indigo", "Ivory", "Khaki", "Lavender", "LavenderBlush", "LawnGreen", "LemonChiffon", "LightBlue", "LightCoral", "LightCyan", "LightGoldenrodYellow", "LightGray", "LightGreen", "LightPink", "LightSalmon", "LightSeaGreen", "LightSkyBlue", "LightSlateGray", "LightSteelBlue", "LightYellow", "Lime", "LimeGreen", "Linen", "Magenta", "Maroon", "MediumAquamarine", "MediumBlue", "MediumOrchid", "MediumPurple", "MediumSeaGreen", "MediumSlateBlue", "MediumSpringGreen", "MediumTurquoise", "MediumVioletRed", "MidnightBlue", "MintCream", "MistyRose", "Moccasin", "NavajoWhite", "Navy", "OldLace", "Olive", "OliveDrab", "Orange", "OrangeRed", "Orchid", "PaleGoldenrod", "PaleGreen", "PaleTurquoise", "PaleVioletRed", "PapayaWhip", "PeachPuff", "Peru", "Pink", "Plum", "PowderBlue", "Purple", "Red", "RosyBrown", "RoyalBlue", "SaddleBrown", "Salmon", "SandyBrown", "SeaGreen", "SeaShell", "Sienna", "Silver", "SkyBlue", "SlateBlue", "SlateGray", "Snow", "SpringGreen", "SteelBlue", "Tan", "Teal", "Thistle", "Tomato", "Turquoise", "Violet", "Wheat", "White", "WhiteSmoke", "Yellow", "YellowGreen")]$Color,
                      $ColorObject,
                      [ValidateSet("none", "select", "lselect")]$Alert,
                      [ValidateSet("none", "colorloop")]$Effect) {
    Process {
        $Service = $(if ($Light.IsGroup) { "groups" } else { "lights" })
        $ServiceState = $(if ($Light.IsGroup) { "action" } else { "state" })
        $Parameters = @{}

        if ($Color -or $ColorObject) { # Turn System.Drawing.Color or KnownColor string into HSB
            if ($Color) {
                add-type -AssemblyName System.Drawing
                $ColorObject = [Drawing.Color]::$Color
            }

            $Hue = [int]($ColorObject.GetHue() * (65535 / 360))
            if (!$Saturation) {
                $Saturation = [int]($ColorObject.GetSaturation() * 255)
            }
            if (!$Brightness) {
                $Brightness = [int]($ColorObject.GetBrightness() * 255)
            }
        }

        if ($Relax) {
            $Hue = 13122
            $Saturation = 211
            $Brightness = 190
        }

        if ($White) {
            $Saturation = 0
            $Brightness = 255
        }

        if ($Default) {
            $Brightness = 255
            $Hue = 14922
            $Saturation = 144
        }

        if ($On) {                    $Parameters.Add('on',$true) } 
        elseif ($Off) {               $Parameters.Add('on',$false) }
        if ($Brightness-ne -1) {      $Parameters.Add('bri',$Brightness) }
        if ($Saturation -ne -1) {     $Parameters.Add('sat',$Saturation) }
        if ($Hue-ne -1) {             $Parameters.Add('hue',$Hue) }
        if ($TransitionTime -ne -1) { $Parameters.Add('transitiontime',$TransitionTime) }
        if ($Alert) {                 $Parameters.Add('alert',$Alert) }
        if ($Effect) {                $Parameters.Add('effect',$Effect) }
        
        # Hack or clever? :|
        if ($AsScheduleCommand) {
            function Invoke-Result { $input }
            function Invoke-RestMethod($Url, $Method, $Body) {
                [pscustomobject]@{
                    'address'=$Url -replace "http://$BridgeIP" #/api/$UserName","/api/0"
                    'method'=$Method.ToUpper()
                    'body'=($Body | ConvertFrom-Json)
                }
            }
        }
        
        Invoke-RestMethod "http://$BridgeIP/api/$UserName/$Service/$($Light.ID)/$ServiceState" -Method Put -Body ([pscustomobject]$Parameters | ConvertTo-Json) | Invoke-Result
    }
}

function Get-HueBridge {
    Invoke-RestMethod "https://www.meethue.com/api/nupnp" | % {
        [pscustomobject]@{'ID'=$_.id
                          'IPAddress'=$_.internalipaddress
                          'MACAddress'=$_.macaddress}
    }
}

function Register-HueUserName($BridgeIP = (gs Hue.BridgeIP), $UserName = (gs Hue.userName)) {
	Invoke-RestMethod "http://$BridgeIP/api" -Method Post -Body (
        [pscustomobject]@{'devicetype'='PowerShell'
                          'username'=$UserName} | ConvertTo-Json) | Invoke-Result
}

function Get-ScreenColor {
    add-type -AssemblyName System.Drawing
add-type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace Colors {
    public class ColorMath {
        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out Rectangle rect);

		public static object GetScreenColor() {
			// Rectangle bounds = Screen.GetBounds(Point.Empty);
            Rectangle bounds = new Rectangle();
            IntPtr handle = GetForegroundWindow();
            GetWindowRect(handle, out bounds);

			using(Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height)) {
				using(Graphics g = Graphics.FromImage(bitmap)) {
					 g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
				}
				return getDominantColor(bitmap);
			}
		}

        public static object getDominantColor(Bitmap bmp) {
            Dictionary<int,int> Hues = new Dictionary<int,int> ();

            for (int x = 0; x < bmp.Width; x++) {
                for (int y = 0; y < bmp.Height; y++) {
                    Color clr = bmp.GetPixel(x, y);

                    int Hue = (int)clr.GetHue();
                    if (Hues.ContainsKey(Hue)) {
                        Hues[Hue]++;
                    } else {
                        Hues[Hue] = 1;
                    }
                }
            }

            var HuesList = new List<KeyValuePair<int, int>>(Hues);
            HuesList.Sort(
                delegate(KeyValuePair<int, int> firstPair, KeyValuePair<int, int> nextPair)
                {
                    return nextPair.Value.CompareTo(firstPair.Value);
                });

            return HuesList[1].Key; // return top color
        }
    }
}
"@ -ReferencedAssemblies System.Drawing, System.Windows.Forms | out-null

    [Colors.ColorMath]::GetScreenColor()
}

function Invoke-Result([Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]$ret) {
    if ($ret.success) {
        # TODO: build a set of expected results and compare with success results
        $ret # debugging
    } elseif ($ret.error.description) {
        write-error -Message ($ret.error.description -replace "`n","")
    } else {
        $ret # dunno, be friendly
    }
}

function Get-HueSchedule($BridgeIP = (gs Hue.BridgeIP), 
                         $UserName = (gs Hue.userName)) {
	$pack = Invoke-RestMethod "http://$BridgeIP/api/$UserName/schedules" -Method Get
    $Schedules = $pack | Get-Member | ? { $_.MemberType -eq "NoteProperty" } |  select -exp Name | % {
            [PSCustomObject]@{'ID'=[int]::Parse($_)
                              'Name'=$pack.$_.Name} 
    }

    $Schedules | % {
        $Sched = Invoke-RestMethod "http://$BridgeIP/api/$UserName/schedules/$($_.ID)"
        [PsCustomObject]@{
            'ID'=$_.ID
            'Name'=$Sched.name
            'Description'=$Sched.description
            'Command'=$Sched.command
            'Time'=$Sched.time
        }
    }
}

function Set-HueSchedule($BridgeIP = (gs Hue.BridgeIP), 
                         $UserName = (gs Hue.userName),
                         $Name = "Unknown Schedule",
                         $Description,
                         $Command,
                         [DateTime]$Time) {
    $Parameters = @{}
    if ($Name) { $Parameters.Add('name',$Name) }
    if ($Description) { $Parameters.Add('description', $Description) }
    if ($Command) { $Parameters.Add('command', $Command) }
    if ($Time) { $Parameters.Add('time', $Time.ToString("yyyy-MM-ddTHH:mm:ss")) }

    ([pscustomobject]$Parameters | ConvertTo-Json) 
	Invoke-RestMethod "http://$BridgeIP/api/$UserName/schedules" -Method Post -Body (([pscustomobject]$Parameters | ConvertTo-Json).Replace(" ","")) | Invoke-Result
}

function Remove-HueSchedule($BridgeIP = (gs Hue.BridgeIP), $UserName = (gs Hue.userName), [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]$Schedule) {
    Process {
	    Invoke-RestMethod "http://$BridgeIP/api/$UserName/schedules/$($Schedule.ID)" -Method Delete | Invoke-Result
    }
}

function Get-HueConfiguration($BridgeIP = (gs Hue.BridgeIP), $UserName = (gs Hue.userName)) {
	Invoke-RestMethod "http://$BridgeIP/api/$UserName/config" -Method Get
}


function Set-HueConfiguration(
                      $BridgeIP = (gs Hue.BridgeIP), 
                      $UserName = (gs Hue.userName),
                      [int]$ProxyPort,
                      [bool]$LinkButton) {
    Process {
        $Parameters = @{}

        if ($ProxyPort) { $Parameters.Add('proxyport',$ProxyPort) }
        if ($Name) { $Parameters.Add('name',$Name) }
        if ($LinkButton) { $Parameters.Add('linkbutton',$LinkButton) }
        # TODO: incomplete
        Invoke-RestMethod "http://$BridgeIP/api/$UserName/config" -Method Put -Body ([pscustomobject]$Parameters | ConvertTo-Json) | Invoke-Result
    }
}

function Get-HueTime {
    # I can't figure out a way to set the time on the bridge, and it's off by a few minutes
    # expose the time so callers can create schedules that start at very specific times (instant)
    [datetime]::Parse((Get-HueConfiguration | select -exp UTC))
}

Export-ModuleMember -Function Get-HueBridge
Export-ModuleMember -Function Register-HueUserName
Export-ModuleMember -Function Get-HueLight 
Export-ModuleMember -Function Set-HueLight
Export-ModuleMember -Function Get-ScreenColor
Export-ModuleMember -Function Get-HueSchedule
Export-ModuleMember -Function Set-HueSchedule
Export-ModuleMember -Function Get-HueConfiguration
Export-ModuleMember -Function Get-HueTime
Export-ModuleMember -Function Remove-HueSchedule
Export-ModuleMember -Function Set-HueConfiguration