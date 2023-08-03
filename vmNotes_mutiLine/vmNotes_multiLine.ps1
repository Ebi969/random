$out = Get-vm SRV005785 | Select Name, @{n="Notes"; e={(($vmBase.Notes).Split("`n")) -join " "}
                               } | Export-Excel .\test.xlsx