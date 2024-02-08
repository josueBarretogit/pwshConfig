Import-Module ImportExcel
Add-Type -AssemblyName PresentationFramework

$VUE = "C:\Users\USUARIO\Desktop\empleadosFrontend" 
$API = "C:\Users\USUARIO\Desktop\empleadosApi" 
$CRINMO = "C:\xampp\htdocs\CRINMO" 


oh-my-posh init pwsh --config 'C:\Users\USUARIO\AppData\Local\Programs\oh-my-posh\themes\clean-detailed.omp.json'  | Invoke-Expression

function SwitchToNvchad  {
  Start-Job -command {

    Remove-Item  "C:\Users\USUARIO\AppData\Local\nvim\*" -Recurse -Force
      Remove-Item  "C:\Users\USUARIO\AppData\Local\nvim-data\*" -Recurse -Force

      Copy-Item "C:\Users\USUARIO\AppData\Local\nvim-nvchad\*" -Destination  "C:\Users\USUARIO\AppData\Local\nvim\" -Recurse
      Copy-Item "C:\Users\USUARIO\AppData\Local\nvim-nvchad-data\*" -Destination  "C:\Users\USUARIO\AppData\Local\nvim-data\" -Recurse
  } -Name "loadginNvchad"

}

function UpdateNvChadConfig {
  Start-Job -command {

    Remove-Item  "C:\Users\USUARIO\AppData\Local\nvim-nvchad\*" -Recurse -Force
      Remove-Item  "C:\Users\USUARIO\AppData\Local\nvim-nvchad-data\*" -Recurse -Force
      Copy-Item "C:\Users\USUARIO\AppData\Local\nvim\*" -Destination "C:\Users\USUARIO\AppData\Local\nvim-nvchad" -Recurse 
      Copy-Item "C:\Users\USUARIO\AppData\Local\nvim-data\*" -Destination "C:\Users\USUARIO\AppData\Local\nvim-nvchad-data" -Recurse 

  } -Name "updatingNvChad"

}

function SwitchToVsCode   {
  Start-Job  -command {
    Remove-Item  "C:\Users\USUARIO\AppData\Local\nvim\*" -Recurse -Force
      Remove-Item  "C:\Users\USUARIO\AppData\Local\nvim-data\*" -Recurse -Force
      Copy-Item "C:\Users\USUARIO\AppData\Local\nvim-vscode\*" -Destination  "C:\Users\USUARIO\AppData\Local\nvim\" -Recurse
      Copy-Item "C:\Users\USUARIO\AppData\Local\nvim-vscode-data\*" -Destination  "C:\Users\USUARIO\AppData\Local\nvim-data\" -Recurse

  } -Name "loadingVscode"
}



function GetBitacora {
  [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,HelpMessage="bitacoraNumber")] 
        [int]$numberBitacora,
        [Parameter(Mandatory=$true,HelpMessage="bitacoraNumber")] 
        [datetime]$fechaInicial
        )
      PROCESS {

        $excelDirectory = "C:\Users\USUARIO\Downloads\excel"
          $namePreviousExcell = "bitacora$($numberBitacora - 1).xlsx"

          $excel = Open-ExcelPackage -Path "$excelDirectory\$namePreviousExcell"

          $dateFromAndToCell = "T10"
          $deliverDate = "U45"
          $bitacoraNumber = "P10"
          $initialDate = "K26"
          $LastDate = "M26"

          $dateFrom = $fechaInicial.ToShortDateString() 
          $dateTo =  $fechaInicial.AddDays(15).ToShortDateString() 

          $longDateFrom = $fechaInicial.ToString("dd, MMMM, yyyy")
          $longDateTo =  $fechaInicial.AddDays(15).ToString("dd, MMMM, yyyy")

          $workSheet = $excel.Workbook.Worksheets[1]

          $workSheet.Cells[$dateFromAndToCell].Value = "$longDateFrom a $longDateTo"
          $workSheet.Cells[$deliverDate].Value = $dateTo
          $workSheet.Cells[$bitacoraNumber].Value = "$( $numberBitacora - 1  )"
          $workSheet.Cells[$initialDate].Value = "$dateFrom"
          $workSheet.Cells[$LastDate].Value = "$dateTo"


          Close-ExcelPackage $excel
          Copy-Item "$excelDirectory\$namePreviousExcell" -Destination "$excelDirectory\bitacora$numberBitacora.xlsx"
      } 
  END {
    Write-Output "Termino la bitacora"
  }
}



function generateBitacoras {

  param()

    $date = (Get-Date -Date "25/01/2024")
    $numberBitacora = 8
    Rename-Item "C:\Users\USUARIO\Downloads\excel\bitacora6.xlsx"  -NewName "bitacora7.xlsx"
    for ($1 = 8; $1 -lt 14; $1++) {
      $date = $date.AddDays(15)

        GetBitacora -numberBitacora $numberBitacora   -fechaInicial $date 
        $numberBitacora++
    }
}

function getBitacorasDate {
  param()
    PROCESS {

      $date = (Get-Date -Date "25/01/2024")
        for ($1 = 8; $1 -lt 14; $1++) {
          $date = $date.AddDays(15)
            Write-Output $date.ToLongDateString()
        }
      return $fechas
    }
  END {
  }
}

function ShowAlert {
  param()
    PROCESS {
      $dates = Get-Content -Path "C:\scripts\fechasEntrega.txt"
        $today = (Get-Date).ToLongDateString()
        foreach($date in $dates) {
          if ($today -eq $date) {
            [System.Windows.MessageBox]::Show("Hoy toca subir bitacora", 'Bitacora')
          }
        }
    }
  END {
  }
}
ShowAlert
#super util para buscar roles en crinmo
function SearchForAMatch {
  [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, HelpMessage = "Directorio a buscar")]
        [string]$Path,
        [Parameter(Mandatory = $true, HelpMessage = "Palabra a buscar")]
        [string]$Word
        )
      PROCESS {
        Get-ChildItem $Path -Recurse -File | Where-Object { Select-String -Path $_.FullName -SimpleMatch $Word } | 
        Format-Table -GroupBy directoryName 
      }
      END {}
}
#-SimpleMatch "rol"} | Format-Table -GroupBy directoryname
