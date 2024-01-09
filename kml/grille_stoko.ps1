# vim: set ft=ps1 tabstop=2 shiftwidth=2 expandtab :
# Important» :set bomb
# & C:\coffee\dev\kml\grille_stoko.ps1

Set-Variable -Name ROOT_SCRIPT -Option Constant `
  -Value (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)

Set-Variable -Name SAVED_PATH -Option Constant `
  -Value $Env:PATH

Push-Location `
  -LiteralPath $ROOT_SCRIPT `
  -ErrorAction SilentlyContinue
If ( $? -cne $TRUE ) { Return }

Try {
  Set-StrictMode -Version Latest

  If (-not ('ZXing.BarcodeWriter' -as [type])) {
    Add-Type -Path 'zxing.dll'
  }
  If (Test-Path -LiteralPath 'tmp' -PathType Container) {
    Remove-Item -LiteralPath 'tmp' -Recurse -Force -ErrorAction SilentlyContinue
  }
  If (-not (Test-Path -LiteralPath 'tmp' -PathType Container)) {
    New-Item -Name 'tmp' -ItemType Directory | Out-Null
  }

  $BarcodeWriter = New-Object -TypeName ZXing.BarcodeWriter
  $NbrPages = 10
  $NbrDuplicates = 2

  $Density = 200

  [ScriptBlock]$F = {
    Param ([Boolean]$ForceGeneration)

    If (($ForceGeneration -ceq $TRUE) -or ($i % 21 -ceq 0)) {
      & 'magick' (
        @(
          'montage'
          '-units', 'PixelsPerCentimeter',
          '-density', $Density.ToString('0')
        ) +
        @(
          1..(& { While (21 -lt $i) { $i = $i - 21 } Return $i }) | % {
            ('tmp/output-{0:d2}.png' -f $_)
          }
        ) +
        @(
          '-tile', '3x7',
          '-geometry', '+0+0',
          '-quality', '100',
          'tmp/grid.png'
        )
      )
      & 'magick' @(
        'convert',
        '-units', 'PixelsPerCentimeter',
        '-density', $Density.ToString('0'),
        '-size', ((21 * $Density).ToString('0'), (29.7 * $Density).ToString('0') -join 'x'),
        'xc:white',
        'tmp/grid.png',
        '-gravity', 'NorthWest',
        '-geometry', ('+', (1.5 * $Density).ToString('0'), '+', (0.85 * $Density).ToString('0')),
        '-composite',
        ('tmp/page-{0:d4}.png' -f ++$PageNo)
      )
    }
    If (($ForceGeneration -ceq $TRUE) -or ($i % (21 * $NbrPages) -ceq 0)) {
      $StartX = ($PageNo-1) - (($PageNo-1) % $NbrPages) + 1
      $Arguments = (
        @(
          '-density', $Density
        ) +
        @(
          $StartX..$PageNo | % {
            ('tmp/page-{0:d4}.png' -f $_)
          }
        ) +
        @(
          ('stoko-7x3-de-page-{0:d4}-à-page-{1:d4}.pdf' -f $StartX, $PageNo)
        )
      )
      (@('magick') + $Arguments -join ' ') | Out-Host
      & 'magick' $Arguments
      $StartX..$PageNo | % {
        Remove-Item -LiteralPath ('tmp/page-{0:d4}.png' -f $_) -Force
      }
    }
  }

  [ScriptBlock]$ProcessStr = {
    Param ([String]$s)
    $s = (
      [regex]::Replace(
        $s, '\s+', ' ',
          [Text.RegularExpressions.RegexOptions]::CultureInvariant -bor
            [Text.RegularExpressions.RegexOptions]::IgnoreCase
      )
    )
    $s = $s.Replace('%', '%%') 
    Return $s
  }

  $CSV = Import-Csv -LiteralPath .\grille_stoko.csv -Delimiter "`t" -Encoding utf8

  $CSV = (
    $CSV | Sort-Object -Property @{
      e = {
        $_.'Code Rayon'.Trim().ToLower([CultureInfo]::InvariantCulture)
      }
    }
  )

  $CSV | % { [Int32]$PageNo = 0; [Int32]$i = 0 } {

#$_ | out-host

    <#
    If ($i -lt 7*3*62) {
      If (++$i % 21) {
        ++$PageNo
      }
      Return
    }
    #>

    $BarcodeWriter.Format = [ZXing.BarcodeFormat]::CODE_128

    $BarcodeWriter.Options = New-Object ZXing.Common.EncodingOptions
    $BarcodeWriter.Options.Width = 1440
    $BarcodeWriter.Options.Height = 720
    $BarcodeWriter.Options.NoPadding = $TRUE
    $BarcodeWriter.Options.PureBarcode = $TRUE

    $BarcodeBitmap = $BarcodeWriter.Write($_.'Référence'.Trim())

    $BarcodeBitmap.Save([IO.Path]::Combine($PWD.Path, 'tmp', ('barcode-{0:d2}.png' -f ($i % 21 + 1))))

    $magickArgs = @(
        'convert',
        '-respect-parentheses'
        '-units', 'PixelsPerCentimeter',
        '-density', $Density.ToString('0'),

        '(',
          '-size', ((6 * $Density).ToString('0'), '' -join ''),
          '-gravity', 'center',
          '(',
            ('tmp/barcode-{0:d2}.png' -f ($i % 21 + 1)),
            '-trim',
            '-resize', ((2.5 * $Density).ToString('0'), '' -join ''),
          ')',
          '(',
            '-font', 'Tahoma',
            '-pointsize', '22',
            ('caption:' + $_.'Référence'.Trim()),
          ')',
          '-append',
          '-trim',
          '-write', 'mpr:barcode',
          '+delete',
        ')',

        '(',
          '-background', 'transparent',
          '-size', ((6 * $Density * 0.95).ToString('0'), '' -join ''),
          '-gravity', 'West',
          '-font', 'Tahoma-Bold',
          '-pointsize', '28',
          ('caption:' + (& $ProcessStr $_.'Désignation'.Trim())),

          '(',
            '+clone',
            '-background', 'white',
            '-alpha', 'background',
            '-channel', 'A',
            '-blur', ((0.1 * $Density).ToString('0'), (0.1 * $Density) -join 'x'),
            '-level', '0,0%',
          ')',

          '-composite',

          '-gravity', 'North',
          '-crop', ((6 * $Density * 0.95).ToString('0'), 'x', (1.95 * $Density), '+', '0', '+', '0' -join ''),

          '-write', 'mpr:desc',
          '+delete',
        ')',

        '-size', ((6 * $Density).ToString('0'), (4 * $Density) -join 'x'),
        'xc:#FFF',

        '-font', 'Tahoma',
        '-gravity', 'SouthWest',
        '-pointsize', '22',
        '-annotate', ('0', 'x', '0', '+', (0.07 * $Density).ToString('0'), '+', (0.07 * $Density).ToString('0') -join ''),
        ('' + (& $ProcessStr $_.'Code Rayon'.Trim())),

        'mpr:barcode',
        '-geometry', ('+', (0.37 * $Density).ToString('0'), '+', (0.13 * $Density).ToString('0') -join ''),
        '-gravity', 'SouthEast',
        '-composite',

        '-font', 'Tahoma-Bold',
        '-gravity', 'North',
        '-pointsize', '48',
        '-annotate', ('0', 'x', '0', '+', '0', '+', (0.3 * $Density).ToString('0') -join ''),
        ('' + (& $ProcessStr $_.'PV MB'.Trim())),

        'mpr:desc',
        '-gravity', 'center',
        '-composite',

        #'-bordercolor', 'black',
        #'-shave', '10',
        #'-border', '10',

        ('tmp/output-{0:d2}.png' -f ($i % 21 + 1))
    )

    #(@('magick') + $magickArgs -join ' ') | Out-Host

    & 'magick' $magickArgs

    If ($NbrDuplicates -cne 0) {
      1..$NbrDuplicates | % {
        [Int32]$PreviousNo = ($i % 21 + 1)
        ++$i
        . $F $FALSE
        Copy-Item `
          -LiteralPath ('tmp/output-{0:d2}.png' -f $PreviousNo) `
          -Destination ('tmp/output-{0:d2}.png' -f ($i % 21 + 1))
      }
    }
    ++$i
    . $F $FALSE
  }
  If (($PageNo + 1) % $NbrPages -cne 0) {
    . $F $TRUE
  }
}
Finally {
  $Env:PATH = $SAVED_PATH

  Pop-Location
}

