# tworzenie interfejsu
cls

$DebugPreference = "Continue"

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Net.Mail")

$button_width = 80
$button_height = 70

function CreateButton {
    param
    (
        [string]$color,
        [scriptblock]$callback,
        [string]$text,
        [string]$tooltip
    )
	
    # tworzy przycisk z obramowaniem o zadanym $color. $text, $tooltip i $callback 

    $form = New-Object System.Windows.Forms.Panel 
    $form.Size = New-Object System.Drawing.Size(($button_width + 10), ($button_height + 10))
    if ($color -ne "def")
    { $form.BackColor = $color }
    $form.Margin = 0
    $form.BorderStyle = [System.Windows.Forms.BorderStyle]::None
			
    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Size(5, 5)
    $button.Size = New-Object System.Drawing.Size($button_width, $button_height)
    $button.Text = $text
    $button.Margin = 0
    $button.BackColor = "white"
	
    $tt = New-Object System.Windows.Forms.ToolTip
    $tt.SetToolTip($button, $tooltip)
	
    $button.Add_Click($callback)
    $form.Controls.Add($button)
	
    return $form
}

function CreateRowGrid {
    param
    (
        [int]$columns,
        [bool]$setColumnStype = $True
    )
	
    # tworzy TableLayoutPanel z 1 wierzem i $columns liczba kolumn.
    # Zwraca TableLayoutPanel
	
    $table = New-Object System.Windows.Forms.TableLayoutPanel
    $table.Dock = "Fill"
    $table.AutoSize = $True
    $table.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $table.RowCount = 1
    $table.Margin = 0
    $table.ColumnCount = $columns
	
    if ($setColumnStype) {
        foreach ($col_num in 1 .. $columns) {
            $cs = New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)
            $dummy = $table.ColumnStyles.Add($cs)
        }
    }
    return $table
}

function CreateLabel {
    param
    (
        [string]$text
    )
	
    # tworzy etykiete z $text
	
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.AutoSize = $True
    $label.Dock = "Fill"
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
		
    return $label
}

# fukcje operacyjne

# oper[0] - pierwszy argument, oper[1] - drugi argument, oper[2] - typ operacji

$oper = @($null, $null, $null)


function clearE {
    if ($oper[0] -eq $null)
    { return 0 }
    else
    { $oper[1] = $null }
    $disp.Text = "0"
}

function clearC {

    param (
        [bool]$clear = $True
    )
    $oper[0] = $null
    $oper[1] = $null
    $oper[2] = $null
    if ($clear)
    { $disp.Text = "0" }
}

function clearCh {

    if ($disp.Text -ne "0") {
        if ($disp.Text.Length -eq 1 -or ($disp.text.StartsWith("-") -and $disp.text.Length -eq 2))
        { $disp.Text = "0" }
        else
        { $disp.text = $disp.Text.Remove($disp.Text.Length - 1, 1) }
        if ($disp.text.EndsWith("."))
        { $disp.text = $disp.Text.Remove($disp.Text.Length - 1, 1) }
        $disp.Update()
    }

}

function addDisp {
    param(
        [string]$text
    )

    if ($text -eq ".") {
        if ($disp.Text.Contains("."))
        { return 0 }
        else
        { $disp.Text = $text.Insert(0, $disp.Text) }
    }
    else {
        if ($disp.Text -eq 0)
        { $disp.Text = $text }
        else
        { $disp.Text = $text.Insert(0, $disp.Text) }
    }

}

function changeSign {
    if ($disp.Text -eq "0") {
        Write-Debug "Nie zmieniam znaku dla zera!"
        return 0 
    }
    if ($disp.Text.StartsWith("-")) {
        Write-Debug "Zmieniam znak dla ujemnej"
        $disp.Text = $disp.Text.Remove(0, 1) 
    }
    else { 
        Write-Debug "Zmieniam znak dla dodatniej"
        $disp.Text = $disp.Text.Insert(0, "-") 
    }
}

function divide {
    $oper[0] = [double] $disp.Text
    $oper[2] = "/"
    $disp.Text = "0"
    Write-Debug "divide ready"
}

function multiply {
    $oper[0] = [double] $disp.Text
    $oper[2] = "*"
    $disp.Text = "0"
    Write-Debug "multiply ready"
}

function addition {
    $oper[0] = [double] $disp.Text
    $oper[2] = "+"
    $disp.Text = "0"
    Write-Debug "addition ready"
}

function substract {
    $oper[0] = [double] $disp.Text
    $oper[2] = "-"
    $disp.Text = "0"
    Write-Debug "substraction ready"
}

function execute {
    Write-Debug "executing"
    if ($oper[2] -ne $null) {
        Write-Debug "function is selected"
        $oper[1] = [double] $disp.Text
        switch ($oper[2]) {
            "/" {
                Write-Debug "executing division"
                if ($oper[1] -eq 0) {
                    [System.Windows.MessageBox]::Show('Nie dziel przez zero!')
                }
                else {
                    $disp.Text = [string]($oper[0] / $oper[1])
                    $disp.Update()
                }
            }

            "*" {
                Write-Debug "executing multiplication"
                $disp.Text = [string]($oper[0] * $oper[1])
                $disp.Update()
            }
            "+" {
                Write-Debug "executing addition"
                $disp.Text = [string]($oper[0] + $oper[1])
                $disp.Update()
            }
            "-" {
                Write-Debug "executing substraction"
                $disp.Text = [string]($oper[0] - $oper[1])
                $disp.Update()
            }


        }
    }
    else { 
        [System.Windows.MessageBox]::Show('Najpierw wybierz operację!')
    }
    #$disp.Text = [string]$disp.Text.Replace(".",",")
    clearC $False
}


# Main Form
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Kalkulator"
$objForm.Size = New-Object System.Drawing.Size(400, 520) 
$objForm.Opacity = 1
$objForm.BackColor = "white"
$objForm.KeyPreview = $True
$objForm.MaximizeBox = $False
$objForm.Add_KeyDown( { if ($_.KeyCode -eq "Escape") 
        { $objForm.Close() } })



# ROW 0
# dodajemy wiesze do mainGrid
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)
#$dummy = $mainGrid.RowStyles.Add($rs)
# ROW 1
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)
#$dummy = $mainGrid.RowStyles.Add($rs)
# ROW 2
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)
#$dummy = $mainGrid.RowStyles.Add($rs)
# ROW 3
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)
#$dummy = $mainGrid.RowStyles.Add($rs)
# ROW 4
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)
$dummy = $mainGrid.RowStyles.Add($rs)
## ROW 5
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)
#$dummy = $mainGrid.RowStyles.Add($rs)
# ROW 6
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)
#$dummy = $mainGrid.RowStyles.Add($rs)

$mainGrid = New-Object System.Windows.Forms.TableLayoutPanel
$mainGrid.Dock = "Fill"
$mainGrid.AutoSize = $True
$mainGrid.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$mainGrid.RowCount = 6
$mainGrid.ColumnCount = 1
$mainGrid.Margin = 0
$objForm.Controls.Add($mainGrid)

# ROW 1 - wyświetlacz
$disp = New-Object System.Windows.Forms.TextBox
$disp.ReadOnly = $True
$disp.Dock = "Fill"
$disp.AutoSize = $True
$disp.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor `
    [System.Windows.Forms.AnchorStyles]::Top    -bor `
    [System.Windows.Forms.AnchorStyles]::Left   -bor `
    [System.Windows.Forms.AnchorStyles]::Right
$disp.Size = New-Object System.Drawing.Size(260, 60) 
$disp.Text = "0"
$disp.Multiline = $False
$disp.MaxLength = 19
$disp.Font = New-Object System.Drawing.Font("Consolas", "24")
$disp.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$disp.ForeColor = "Black"
$disp.BackColor = "White"
$disp.ReadOnly = $True
$disp.TextAlign = "Right"
$mainGrid.Controls.Add($disp)

# ROW 2 - klawisze
$r2 = CreateRowGrid 4
$mainGrid.Controls.Add($r2, 0, 1)
$b20 = CreateButton "def" { clearE }  `
    "CE" `
    "czyści ekran"
$r2.Controls.Add($b20, 0, 1)

$b21 = CreateButton "def" { clearC }  `
    "C" `
    "czyści pamięć"
$r2.Controls.Add($b21, 1, 1)

$b22 = CreateButton "def" { clearCh }  `
    "←" `
    "czyści znak"
$r2.Controls.Add($b22, 2, 1)

$b23 = CreateButton "def" { divide }  `
    "÷" `
    "dzielenie"
$r2.Controls.Add($b23, 3, 1)

# ROW 3 - 

$r3 = CreateRowGrid 4
$mainGrid.Controls.Add($r3, 0, 1)
$b30 = CreateButton "def" { addDisp 7 }  `
    "7" `
    "7"
$r2.Controls.Add($b30, 0, 2)

$b31 = CreateButton "def" { addDisp 8 }  `
    "8" `
    "8"
$r2.Controls.Add($b31, 1, 2)

$b32 = CreateButton "def" { addDisp 9 }  `
    "9" `
    "9"
$r2.Controls.Add($b32, 2, 2)

$b33 = CreateButton "def" { multiply }  `
    "х" `
    "mnożenie"
$r2.Controls.Add($b33, 3, 2)

# ROW 4 - 
$r4 = CreateRowGrid 4
$mainGrid.Controls.Add($r4, 0, 1)
$b40 = CreateButton "def" { addDisp 4 }  `
    "4" `
    "4"
$r2.Controls.Add($b40, 0, 3)

$b41 = CreateButton "def" { addDisp 5 }  `
    "5" `
    "5"
$r2.Controls.Add($b41, 1, 3)

$b42 = CreateButton "def" { addDisp 6 }  `
    "6" `
    "6"
$r2.Controls.Add($b42, 2, 3)

$b43 = CreateButton "def" { substract }  `
    "-" `
    "odejmowanie"
$r2.Controls.Add($b43, 3, 3)


# ROW 5 - 
$r5 = CreateRowGrid 4
$mainGrid.Controls.Add($r5, 0, 1)
$b50 = CreateButton "def" { addDisp 1 } "1" "1"
$r2.Controls.Add($b50, 0, 4)

$b51 = CreateButton "def" { addDisp 2 } "2" "2"
$r2.Controls.Add($b51, 1, 4)

$b52 = CreateButton "def" { addDisp 3 } "3" "3"
$r2.Controls.Add($b52, 2, 4)

$b53 = CreateButton "def" { addition } "+" "dodawanie"
$r2.Controls.Add($b53, 3, 4)

# ROW 6 - 	
$r6 = CreateRowGrid 4
$mainGrid.Controls.Add($r6, 0, 1)
$b60 = CreateButton "def" { changeSign } "±" "zmiana znaku"
$r2.Controls.Add($b60, 0, 5)

$b61 = CreateButton "def" { addDisp 0 } "0" "0"
$r2.Controls.Add($b61, 1, 5)

$b62 = CreateButton "def" { addDisp "." } "." "."
$r2.Controls.Add($b62, 2, 5)

$b63 = CreateButton "def" { execute } "=" "wykonaj"
$r2.Controls.Add($b63, 3, 5)
	
# pokaz formularz

$objForm.Topmost = $False

$objForm.Add_Shown( { $objForm.Activate() })
[void] $objForm.ShowDialog()