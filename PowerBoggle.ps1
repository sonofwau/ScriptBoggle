<# Patch notes for v04.10

Updated the word search to highlight found words.

#>
#region Functions

# Init creates all script wide variables (except for those used by FindWords)
Function Init () {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, System.Windows.Forms
    cls
    $script:Seconds = 240
    $script:DebugSeconds = 20
    $script:CountDown = 3

    $script:PSPath = Split-Path $SCRIPT:MyInvocation.MyCommand.Path
    $script:WordFile = "$PSPath/words.txt"
    $Script:Words = [ordered]@{}    
    $script:timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]"0:0:1"
    $Script:foundWords = $false
    $Script:boardCombos = $false
    $script:BoardSize = 6
    $script:debuggingBoard = @{}
    $script:letterCount = @(0,0)
    $script:topFive = [ordered]@{0="";1="";2="";3="";4=""}
    [System.Collections.ArrayList]$script:posList = @("noun","pronoun","verb","adjective","adverb","preposition","conjunction","interjection")

    $script:Window = [System.Windows.Window]@{Title="PowerBoggle"; Width=660; Height=660}
    $script:VB = [System.Windows.Controls.Viewbox]@{}
    $script:Window.AddChild($VB)
    $script:Grid = [System.Windows.Controls.Grid]@{ShowGridLines=0;MinHeight=660;MinWidth=660}
    $script:VB.AddChild($Grid)

    $script:Grid.RowDefinitions.Add([System.Windows.Controls.RowDefinition]@{Height="*"})
    $script:Grid.RowDefinitions.Add([System.Windows.Controls.RowDefinition]@{Height="*"})
    $script:StartButton = [System.Windows.Controls.Button]@{Content="Start";FontSize=50;Background="LightGreen"}
    $script:GameSelect_Combo = [System.Windows.Controls.ComboBox]@{VerticalContentAlignment="Center";HorizontalContentAlignment = "Center";FontSize=50;ItemsSource=@("Classic","New","Big_Original","Big_Challenge","Big_Deluxe","Big_2012","Super_Big","Debugging_Mode");SelectedValue="Super_Big"}
    $script:countdownWords = @{3 = @{"C10" = "G"; "C11" = "e"; "C12" = "t"; "C21" = "R"; "C22" = "e"; "C23" = "a"; "C24" = "d"; "C25" = "y";};	2 = @{"C10" = "G"; "C11" = "e"; "C12" = "t"; "C21" = "S"; "C22" = "e"; "C23" = "t";}; 1 = @{"C20" = "B"; "C21" = "O"; "C22" = "G"; "C23" = "G"; "C24" = "L"; "C25" = "E";}}
    Add-ToGrid -Control $GameSelect_Combo -Parent $Grid -X 0 -Y 0 -ZIndex 1 -RowSpan 1 -ColumnSpan 1
    Add-ToGrid -Control $StartButton -Parent $Grid -X 0 -Y 1 -ZIndex 1 -RowSpan 1 -ColumnSpan 1
}

Function Word-Search () {
    Update-Definition -word $searchBox.Text
    Find-IndividualWords -word $searchBox.Text -boardCombos $boardCombos -gridSize 6 -hashedBoard $boggleBoard
    if ($script:words.contains($searchBox.Text)) {
        Highlight-Words -word $searchBox.Text
    } else {
        $Grid.Children | %{if($_.GetType().Name -eq "Border"){$_.Background = "Black";$_.Child.Foreground = "White";$_.Child.Background = "Black"}}
    }
    $listbox.ItemsSource = $script:Words.keys
}

Function Find-Stuff () {
    param ($element) 
    $parent = $element.parentelement

    if ($parent.className -eq "vg") {
        return $null
    }

    foreach ($child in $parent.children) {
        if ($child.classname -eq "lb badge mw-badge-gray-100 text-start text-wrap d-inline") {
            $found = $child
        }
    }

    if ($found) {
        return $found.innerText
    }

    return Find-Stuff -element $parent
}

Function Definition-Search () {
    param ($word)

    try {$R = iwr -Uri "https://www.merriam-webster.com/dictionary/$word" -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36 Edg/129.0.0.0"
    } catch {return "Error finding $word"}

    $R = [regex]::Split($R.Content, "<div id=`".*-entry-.`" class=`"entry-word-section-container`">")
    
    $root = [regex]::Match($R[0],"(?<=<title>)[a-zA-Z]*(?= Definition &amp; Meaning - Merriam-Webster</title>)").value

    $definition = @()
    $rootRun = New-Object System.Windows.Documents.Run -ArgumentList ("$root has " + ($R.count - 1) + " possible definitions:")
    $rootRun.FontWeight = "Bold"
    $rootRun.FontSize = 18
    $definition += $rootRun

    for ($i = 1; $i -lt $R.count; $i++) {
        $section = ([regex]::Split($R[$i],"class=`"content-section-header`""))[0]
        $html = New-Object -ComObject "HTMLFile"
        $html.IHTMLDocument2_write($section)
        $currentDef = $false

        try {
            $rootWord = ($html.documentElement.getElementsByClassName("cxt") | foreach {$_.innerText}).trim(" ()0123456789:")
        } catch {
            try { 
                $rootWord = ($html.documentElement.getElementsByClassName("ucxt") | foreach {$_.innerText}).trim(" ()0123456789:") 
            } catch {}
        }

        if ($rootWord) {
            $varRun = New-Object System.Windows.Documents.Run -ArgumentList "`n $i) $word is a variant of $rootword"
            $varRun.FontWeight = "Bold"
            $varRun.FontSize = 18
            $definition += $varRun
            continue
        }

        [System.Collections.ArrayList]$varList = @()
        $html.documentElement.getElementsByClassName("if") | foreach {$varList.add($_.innerText) | out-null}
        $html.documentElement.getElementsByClassName("fw-bold ure") | foreach {$varList.add($_.innerText) | out-null}
        try {
            $partOfSpeech = (($html.documentElement.getElementsByClassName("parts-of-speech"))[0].innerText).trim(" ()0123456789")
        } catch {
            continue
        }

        $defrun = New-Object System.Windows.Documents.Run -ArgumentList ("`n $i) "+ ($html.documentElement.getElementsByClassName("hword"))[0].innerText)
        $defrun.FontWeight = "Bold"
        $defRun.FontSize = 18
        $definition += $defRun

        foreach ($PoS in $posList) {
            if ($partOfSpeech -match $PoS) {
                $posText += " ($partOfSpeech"
                $currentDef = $true
                break
            }
        }

        if ($varList) {
            $posText += "; " + ($varList | foreach {$_+","})
            $posText = $posText.Trim(",")
        } 

        $posText += ")"

        $posRun = New-Object System.Windows.Documents.Run -ArgumentList $posText
        $posRun.FontStyle = "Italic"
        $posRun.FontSize = 18
        $definition += $posRun

        if ($currentDef) {
            [System.Collections.ArrayList]$defList = @()
            $html.documentElement.getElementsByClassName("dtText") | foreach {$defList.add($_) | out-null}
            $defText =""
            foreach ($def in $defList) {
                $defText += "`n    -"
                $extra = Find-Stuff -element $def
                if ($extra) {
                    $defText += "$extra " + $def.innerText
                } else {
                    $defText += "" + $def.innerText.trimStart("`: ")
                }
            }
            $defTextRun = New-Object System.Windows.Documents.Run -ArgumentList $defText
            $defTextRun.FontSize = 14
            $definition += $defTextRun
        }
    }

    return $definition
}

Function New-BoggleGrid{
    param(
        [validateset("Classic","New","Big_Original","Big_Challenge","Big_Deluxe","Big_2012","Super_Big","Debugging_Mode")]$GameSelection
    )
    # Return the debugging board
    if ($GameSelection -eq "Debugging_mode") {
        Build-DebuggingBoard
        return $debuggingBoard
    }

    # Return a dictionary of grid positions to random letters
    $Game = @{}
    $Game.Classic = @("AACIOT","ABILTY",@("A","B","J","M","O","Qu"),"ACDEMP","ACELRS","ADENVZ","AHMORS","BIFORX","DENOSW","DKNOTU","EEFHIY","EGKLUY","EGINTV","EHINPS","ELPSTU","GILRUW")
    $Game.New = @('AAEEGN', 'ABBJOO', 'ACHOPS', 'AFFKPS', 'AOOTTW', 'CIMOTU', 'DEILRX', 'DELRVY', 'DISTTY', 'EEGHNW', 'EEINSU', 'EHRTVW', 'EIOSST', 'ELRTTY', @("H","I","M","N","U","Qu"), 'HLNNRZ')
    $Game.Big_Original = @('AAAFRS', 'AAEEEE', 'AAFIRS', 'ADENNN', 'AEEEEM', 'AEEGMU', 'AEGMNN', 'AFIRSY', @("B","J","K","Qu","X","Z"), 'CCENST', 'CEIILT', 'CEIPST', 'DDHNOT', 'DHHLOR', 'DHHLOR', 'DHLNOR', 'EIIITT', 'CEILPT', 'EMOTTT', 'ENSSSU', 'FIPRSY', 'GORRVW', 'IPRRRY', 'NOOTUW', 'OOOTTU')
    $Game.Big_Challenge = @('AAAFRS', 'AAEEEE', 'AAFIRS', 'ADENNN', 'AEEEEM', 'AEEGMU', 'AEGMNN', 'AFIRSY', @("B","J","K","Qu","X","Z"), 'CCENST', 'CEIILT', 'CEIPST', 'DDHNOT', 'DHHLOR', @("I","K","L","M","Qu","U"), 'DHLNOR', 'EIIITT', 'CEILPT', 'EMOTTT', 'ENSSSU', 'FIPRSY', 'GORRVW', 'IPRRRY', 'NOOTUW', 'OOOTTU')
    $Game.Big_Deluxe = @('AAAFRS', 'AAEEEE', 'AAFIRS', 'ADENNN', 'AEEEEM', 'AEEGMU', 'AEGMNN', 'AFIRSY', @("B","J","K","Qu","X","Z"), 'CCNSTW', 'CEIILT', 'CEIPST', 'DDLNOR', 'DHHLOR', 'DHHNOT', 'DHLNOR', 'EIIITT', 'CEILPT', 'EMOTTT', 'ENSSSU', 'FIPRSY', 'GORRVW', 'HIPRRY', 'NOOTUW', 'OOOTTU')
    $Game.Big_2012 = @('AAAFRS', 'AAEEEE', 'AAFIRS', 'ADENNN', 'AEEEEM', 'AEEGMU', 'AEGMNN', 'AFIRSY', 'BBJKXZ', 'CCENST', 'EIILST', 'CEIPST', 'DDHNOT', 'DHHLOR', 'DHHNOW', 'DHLNOR', 'EIIITT', 'EILPST', 'EMOTTT', 'ENSSSU', @("Qu","In","Th","Er","He","An"), 'GORRVW', 'IPRSYY', 'NOOTUW', 'OOOTTU')
    $Game.Super_Big = @("AAAFRS","AAEEEE","AAEEOO","AAFIRS","ABDEIO","ADENNN","AEEEEM","AEEGMU","AEGMNN","AEILMN","AEINOU","AFIRSY",@("An","Er","He","In","Qu","Th"),"BBJKXZ","CCENST","CDDLNN","CEIITT","CEIPST","CFGNUY","DDHNOT","DHHLOR","DHHNOW","DHLNOR","EHILRS","EIILST","EILPST",@("E","I","O","[]","[]","[]"),"EMTTTO","ENSSSU","GORRVW","HIRSTV","HOPRST","IPRSYY",@("J","K","Qu","W","X","Z"),"NOOTUW","OOOTTU")
    
    $diceSet = [System.Collections.ArrayList]::new()
    $diceSet.addRange(@($Game.$GameSelection))
    
    $Board = @{}
    for ($x = 0; $x -lt $script:BoardSize; $x++) {
        $board.add($x, @{})
        for ($y = 0; $y -lt $script:BoardSize; $y++) {
        $rDice = Get-Random -Maximum $diceSet.Count -Minimum 0
        $rSide = Get-Random -Maximum 6 -Minimum 0
        $board[$x].add($y, $diceSet[$rDice][$rSide])
        $diceSet.removeAt($rDice)
        }
    }

    return $Board
}

Function Highlight-Words () {
    param ($word)

    $Grid.Children | %{if($_.GetType().Name -eq "Border"){$_.Background = "Black";$_.Child.Foreground = "White";$_.Child.Background = "Black"}}
    $coordList = @()
    foreach ($coord in $script:words.$word.keys) {
        $coordList += "c" + $script:words.$word.$coord[0] + $script:words.$word.$coord[1]
        foreach ($coord in $coordList) {
            foreach ($square in $grid.Children) {
                if ($square.Name -eq $coord) {
                    $square.Child.Foreground = "Red"
                }
            }
        }
    }
}

Function Update-Definition () {
    param ($word)

    $searchText.Inlines.Clear()

    foreach ($line in (Definition-Search -word $word)) {
        $searchText.Inlines.Add($line)
    }
}

Function Print-Board($BoggleBoard){
    $Dimension = $BoggleBoard.Count
    Write-Host ([System.Environment]::NewLine + "This is a copy of the board for debugging purposes:")
    foreach($x in 0..($Dimension-1)){
        $xLine = "    `$debuggingBoard.add($x, @{"
        foreach($y in 0..($Dimension-1)){
            $xline += "$y=`"" + $BoggleBoard.$x.$y +"`"; "
        }
        Write-Host ($xLine.trimEnd(" ") + "})")
    }
    Write-Host -Object ([System.Environment]::NewLine)
}

# Used for sorting the top5 longest words
Function Check-WordLength () {
    param ($word)

    $holder = $null

    for ($i = 0; $i -lt 5; $i++) {
        if ($word.length -gt $script:topFive[$i].length -or $holder.length -ge $script:topFive[$i].length) {
            $holder = $script:topFive[$i]
            $script:topFive[$i] = $word
            $word = $holder
        }
    }
    return
}

# Adds letters to the grid
Function Add-ToGrid ($Control, $Parent, $X=1, $Y=1, $ZIndex, $RowSpan, $ColumnSpan){
    [System.Windows.Controls.Grid]::SetRow($Control,$Y)
    [System.Windows.Controls.Grid]::SetColumn($Control,$X)
    if($ZIndex){[System.Windows.Controls.Grid]::SetZIndex($Control, $ZIndex)}
    if($RowSpan){[System.Windows.Controls.Grid]::SetRowSpan($Control, $RowSpan)}
    if($ColumnSpan){[System.Windows.Controls.Grid]::SetColumnSpan($Control, $ColumnSpan)}
    $Parent.AddChild($Control)
}

# Builds a custom board for debugging
Function Build-DebuggingBoard () {
    $boardSize = 6
    $debuggingBoard.add(0, @{0="N"; 1="Y"; 2="Y"; 3="A"; 4="O"; 5="N";})
    $debuggingBoard.add(1, @{0="R"; 1="W"; 2="Z"; 3="C"; 4="D"; 5="S";})
    $debuggingBoard.add(2, @{0="Y"; 1="H"; 2="L"; 3="T"; 4="L"; 5="Th";})
    $debuggingBoard.add(3, @{0="S"; 1="N"; 2="E"; 3="M"; 4="D"; 5="T";})
    $debuggingBoard.add(4, @{0="O"; 1="X"; 2="A"; 3="C"; 4="[]"; 5="S";})
    $debuggingBoard.add(5, @{0="A"; 1="O"; 2="U"; 3="G"; 4="S"; 5="I";})
}

# Run as a job in the background to find words
$FindWords = {
    param($hashedBoard, $WordFile, $searchTime, $gridSize)

    # Init-Findwords creates script-wide variables
    Function Init-FindWords () {
        $script:findWordsStartTime = Get-Date
        $script:foundWords = @{}
        $script:boardCombos = @{}
        $script:startTime = Get-Date
        $script:boundCount = 0
        $script:foundWordSearchTime = @{}
        $script:unfoundWordSearchTime = @{}
        $script:boardCombos = Possible-Combos -gridSize $gridSize
        $script:wordList = Import-WordList($wordFile)
    }

    # Check-Letters verifies that the letter for a given coord matches the next letter in the word
    Function Check-Letters() {
        param (
            [int]$px,
            [int]$py,
            [string]$word,
            [int]$wordPos
        )

        $nextLetter = $word[$wordPos]
        if ($hashedBoard[$px][$pY].Length -eq 2 -and $wordPos -lt $word.Length-1) {
            #double check
            $letterWordPos = 2
            $nextLetter = $word[$wordPos]+$word[$wordPos+1]
        } else {
            #single check
            $letterWordPos = 1
            $nextLetter = $word[$wordPos]
        }

        #check if we found it
        if ($hashedBoard[$px][$pY] -eq $nextLetter) {
            return $letterWordPos
        } else {
            return $null
        }
    }

    #check next letter
    Function Check-NextPos() {
        param (
        $passedCoords,
        [string]$word, 
        [int]$wordPos, 
        [string]$testWord
        )

	    #Set x and y for the current coords
        $x = $passedCoords[$passedCoords.Count-1][0]
        $y = $passedCoords[$passedCoords.Count-1][1]

        # Check if the word is complete
        if ($word -eq $testWord) {
            return $passedCoords
        }

        #check each near position
        $possibleCombos = @{0=@(-1,-1);1=@(-1,0);2=@(-1,1);3=@(0,-1);4=@(0,1);5=@(1,-1);6=@(1,0);7=@(1,1);}

        for ($i=0; $i -lt $possibleCombos.Count; $i++) {
            $pX = $x + $possibleCombos[$i][0]
            $pY = $y + $possibleCombos[$i][1]

            if ($pX -lt 0 -or $pY -lt 0) {
                continue
            }

            #check for used coordinates
            foreach ($v in $passedCoords.keys) {
                $newCoords = $true
                if ($passedCoords[$v][0] -eq $pX -and $passedCoords[$v][1] -eq $pY) {
                    $newCoords = $false
                    break
                }
            }

            if ($pX -ge 0 -and $pY -ge 0 -and $pX -le 5 -and $pY -le 5 -and $newCoords) {
                $letterCheck = Check-Letters -px $px -py $py -word $word -wordpos $wordPos

                if ($letterCheck) {
                    $newTestWord = $testWord+$hashedBoard[$px][$pY]
                    $newCoordSet = $passedCoords.Clone()
                    $newCoordSet.Add($passedCoords.Count, @($pX,$pY))
                    $newWordPos = $newTestWord.length
                    $nextCheck = Check-NextPos -passedCoords $newCoordSet -word $word -wordPos $newWordPos -testWord $newTestWord
                }

                if ($nextCheck) {
                    return $nextCheck
                }
            }
        }

        return $null
    }

    # Import-WordList imports the word bank and trims it down to only the possible words
    Function Import-WordList ($wordFile) {
        $stepStartTime = Get-Date
        $wordString = Get-Content -Path $wordFile
        $wordList = @{}
        $truncatedList = @{}
        $totalWordCount = 0
        $truncatedWordCount = 0

        $wordClusterArray = $wordString -split "_"
        foreach ($wordClusterString in $wordClusterArray ) {
            try {
                $wordArray = $wordClusterString -split ":"
                $word = $wordArray[0]
                $wordClusterList = $wordArray[1] -split ","
                $wordList.Add($word, $wordClusterList)
                if ($word -eq "" -or $word -eq $null -or $wordClusterList -eq "" -or $wordClusterList -eq $null) {
                    continue
                }
                $totalWordCount += $wordClusterList.Count
            } catch {}
        }
        $wordlist.remove("")

        foreach ($key in $wordList.Keys) {
            foreach ($k in $boardCombos.Keys) {
                if ($k -imatch $key -and $truncatedList.ContainsKey($key) -eq $false) {
                    $value = $wordList.$Key
                    $truncatedList.Add($key, $value)
                    $truncatedWordCount += $value.Count
                }
            }
        }

        Write-Host ([System.Environment]::NewLine + [System.Environment]::NewLine + "It took " + (Get-Date | foreach-object {$_ - $stepStartTime} | ForEach-Object {$_.totalseconds}) + " seconds to pre-filter $totalWordCount words down to $truncatedWordCount.")

        return $truncatedList
    }

    # Possible-Combos searches the board to find all possible combos
    Function Possible-Combos () {
        param ([int]$gridSize)

        #$pathLength = 3

        Function Get-Paths {
            param (
                [int]$x,
                [int]$y,
                [array]$visited,
                [int]$remaining
            )

            if ($remaining -eq 0) {
                $pathObject = [pscustomobject]@{}
                $pathObject | Add-Member -MemberType NoteProperty -Name "count" -Value $visited.Count
                for ($i = 0; $i -lt $visited.Count; $i++) {
                    $pathObject | Add-Member -MemberType NoteProperty -Name ("Node" + ($i + 1)) -Value $visited[$i]
                }
                return @($pathObject)
            }

            $paths = @()
            $moves = @(
                [tuple]::Create($x + 1, $y),    # Right
                [tuple]::Create($x - 1, $y),    # Left
                [tuple]::Create($x, $y + 1),    # Down
                [tuple]::Create($x, $y - 1),    # Up
                [tuple]::Create($x + 1, $y + 1),# Down-Right
                [tuple]::Create($x - 1, $y - 1),# Up-Left
                [tuple]::Create($x + 1, $y - 1),# Up-Right
                [tuple]::Create($x - 1, $y + 1) # Down-Left
            )

            foreach ($move in $moves) {
                $newX = $move.Item1
                $newY = $move.Item2

                if ($newX -ge 0 -and $newY -ge 0 -and $newX -lt $gridSize -and $newY -lt $gridSize) {
                    $newPosition = [tuple]::Create($newX, $newY)
                    if ($visited -notcontains $newPosition) {
                        $newVisited = $visited + $newPosition
                        $paths += Get-Paths -x $newX -y $newY -visited $newVisited -remaining ($remaining - 1)
                    }
                }
            }

            return $paths
        }

        Function Get-AdjacentPaths {
            param (
                [int]$gridSize,
                [int]$pathLength
            )

            $allPaths = @()
            for ($i = 0; $i -lt $gridSize; $i++) {
                for ($j = 0; $j -lt $gridSize; $j++) {
                    $startPos = [tuple]::Create($j, $i)
                    $newPath = Get-Paths -x $j -y $i -visited @($startPos) -remaining ($pathLength - 1)
                    $allPaths += $newPath
                }
            }

            return $allPaths
        }

        $paths = Get-AdjacentPaths -gridSize $gridSize -pathLength 3
        $paths += Get-AdjacentPaths -gridSize $gridSize -pathLength 2

        $comboCount = @{}
        foreach ($set in $paths) {
            $combo = $null
            $comboCoords = @{}

            for  ($i=0;$i -lt $set.count; $i++) {
                $node = "Node" + ($i+1)
                $px = $set.$node.Item1
                $py = $set.$node.Item2
                $combo += $hashedBoard[$px][$py]
                $comboCoords.add($i, @($px,$py))
            }

            if ($combo -imatch "]" -or $combo.length -lt 3) {
                continue
            }

            try {
                $boardCombos.Add($combo, @{0 = $comboCoords})
            } catch {
                $boardCombos.$combo.Add($boardCombos.$combo.Count, $comboCoords)
            }
        }

        return $boardCombos
    }

    # Main is main
    Function Main-FindWords () {
        # Check each set of words in the word list
        foreach ($key in $wordList.Keys) {
            # If the search time has been exceeded, break if true
            if (Get-Date | foreach-object {$_ - $findWordsStartTime} | ForEach-Object {$_.totalseconds -gt $searchTime}) {
                break
            }
            # Check each word in the set
            foreach ($word in $wordList.$key) {
                # If the search time has been exceeded, break if true
                if (Get-Date | foreach-object {$_ - $findWordsStartTime} | ForEach-Object {$_.totalseconds -gt $searchTime}) {
                    break
                }
            
                # Check for duplicate words and skip
                if($foundWords.Keys -contains $word) {
                    continue
                }

                # Incriment the number of words searched
                $boundCount++
            
                # Find each possible combo
                $pMatches = @()
                foreach ($combo in $boardCombos.Keys) {
                    if ($combo.length -le $word.length -and $combo -eq $word.Substring(0,$combo.length)) {
                        $pMatches += $combo
                    }
                }

                foreach ($combo in $pMatches) {
                    # Check if the word was found in a previous combo
                    if ($foundwords.Contains($word)) {
                        break
                    }

                    foreach ($k in $boardCombos.$combo.Keys) {
                        $startTestWord = $combo
                        $startWordPos = $startTestWord.Length
                        $startCoordSet = $boardCombos.$combo.$k
                        $startCheck = Check-NextPos -passedCoords $startCoordSet -word $word -wordPos $startWordPos -testWord $startTestWord

                        if ($startCheck) {
                            $foundWords.Add($word, $startCheck)
                            break
                        }
                    }
                }
            }
        }
        Write-Host ([System.Environment]::NewLine + [System.Environment]::NewLine+ "It took " + (Get-Date | foreach-object {$_ - $findWordsStartTime} | ForEach-Object {$_.totalseconds}) + " seconds to filter through $boundCount words and find " + $foundWords.Count + " words.")
    }

    Init-FindWords
    Main-FindWords
    
    return @($foundWords,$boardCombos)
}

# Find coords for individual words
Function Find-IndividualWords () {
    param($hashedBoard, $word, $boardCombos, $gridSize)

    if ($script:Words.contains($word)) {
        return
    }

    # Init-Findwords creates script-wide variables
    Function Init-FindWords () {
        $findWordsStartTime = Get-Date
        $startTime = Get-Date
        $boundCount = 0
        $foundWordSearchTime = @{}
        $unfoundWordSearchTime = @{}
    }

    # Check-Letters verifies that the letter for a given coord matches the next letter in the word
    Function Check-Letters() {
        param (
            [int]$px,
            [int]$py,
            [string]$word,
            [int]$wordPos
        )

        $nextLetter = $word[$wordPos]
        if ($hashedBoard[$px][$pY].Length -eq 2 -and $wordPos -lt $word.Length-1) {
            #double check
            $letterWordPos = 2
            $nextLetter = $word[$wordPos]+$word[$wordPos+1]
        } else {
            #single check
            $letterWordPos = 1
            $nextLetter = $word[$wordPos]
        }

        #check if we found it
        if ($hashedBoard[$px][$pY] -eq $nextLetter) {
            return $letterWordPos
        } else {
            return $null
        }
    }

    #check next letter
    Function Check-NextPos() {
        param (
        $passedCoords,
        [string]$word, 
        [int]$wordPos, 
        [string]$testWord
        )

	    #Set x and y for the current coords
        $x = $passedCoords[$passedCoords.Count-1][0]
        $y = $passedCoords[$passedCoords.Count-1][1]

        # Check if the word is complete
        if ($word -eq $testWord) {
            return $passedCoords
        }

        #check each near position
        $possibleCombos = @{0=@(-1,-1);1=@(-1,0);2=@(-1,1);3=@(0,-1);4=@(0,1);5=@(1,-1);6=@(1,0);7=@(1,1);}
        
        for ($i=0; $i -lt $possibleCombos.Count; $i++) {
            $pX = $x + $possibleCombos[$i][0]
            $pY = $y + $possibleCombos[$i][1]

            if ($pX -lt 0 -or $pY -lt 0) {
                continue
            }

            #check for used coordinates
            foreach ($v in $passedCoords.keys) {
                $newCoords = $true
                if ($passedCoords[$v][0] -eq $pX -and $passedCoords[$v][1] -eq $pY) {
                    $newCoords = $false
                    break
                }
            }

            if ($pX -ge 0 -and $pY -ge 0 -and $pX -le 5 -and $pY -le 5 -and $newCoords) {
                $letterCheck = Check-Letters -px $px -py $py -word $word -wordpos $wordPos

                if ($letterCheck) {
                    $newTestWord = $testWord+$hashedBoard[$px][$pY]
                    $newCoordSet = $passedCoords.Clone()
                    $newCoordSet.Add($passedCoords.Count, @($pX,$pY))
                    $newWordPos = $newTestWord.length
                    $nextCheck = Check-NextPos -passedCoords $newCoordSet -word $word -wordPos $newWordPos -testWord $newTestWord
                }

                if ($nextCheck) {
                    return $nextCheck
                }
            }
        }

        return $null
    }

    # Possible-Combos searches the board to find all possible combos
    Function Possible-Combos () {
        param ([int]$gridSize)

        #$pathLength = 3

        Function Get-Paths {
            param (
                [int]$x,
                [int]$y,
                [array]$visited,
                [int]$remaining
            )

            if ($remaining -eq 0) {
                $pathObject = [pscustomobject]@{}
                $pathObject | Add-Member -MemberType NoteProperty -Name "count" -Value $visited.Count
                for ($i = 0; $i -lt $visited.Count; $i++) {
                    $pathObject | Add-Member -MemberType NoteProperty -Name ("Node" + ($i + 1)) -Value $visited[$i]
                }
                return @($pathObject)
            }

            $paths = @()
            $moves = @(
                [tuple]::Create($x + 1, $y),    # Right
                [tuple]::Create($x - 1, $y),    # Left
                [tuple]::Create($x, $y + 1),    # Down
                [tuple]::Create($x, $y - 1),    # Up
                [tuple]::Create($x + 1, $y + 1),# Down-Right
                [tuple]::Create($x - 1, $y - 1),# Up-Left
                [tuple]::Create($x + 1, $y - 1),# Up-Right
                [tuple]::Create($x - 1, $y + 1) # Down-Left
            )

            foreach ($move in $moves) {
                $newX = $move.Item1
                $newY = $move.Item2

                if ($newX -ge 0 -and $newY -ge 0 -and $newX -lt $gridSize -and $newY -lt $gridSize) {
                    $newPosition = [tuple]::Create($newX, $newY)
                    if ($visited -notcontains $newPosition) {
                        $newVisited = $visited + $newPosition
                        $paths += Get-Paths -x $newX -y $newY -visited $newVisited -remaining ($remaining - 1)
                    }
                }
            }

            return $paths
        }

        Function Get-AdjacentPaths {
            param (
                [int]$gridSize,
                [int]$pathLength
            )

            $allPaths = @()
            for ($i = 0; $i -lt $gridSize; $i++) {
                for ($j = 0; $j -lt $gridSize; $j++) {
                    $startPos = [tuple]::Create($j, $i)
                    $newPath = Get-Paths -x $j -y $i -visited @($startPos) -remaining ($pathLength - 1)
                    $allPaths += $newPath
                }
            }

            return $allPaths
        }

        $paths = Get-AdjacentPaths -gridSize $gridSize -pathLength 3
        $paths += Get-AdjacentPaths -gridSize $gridSize -pathLength 2

        $comboCount = @{}
        foreach ($set in $paths) {
            $combo = $null
            $comboCoords = @{}

            for  ($i=0;$i -lt $set.count; $i++) {
                $node = "Node" + ($i+1)
                $px = $set.$node.Item1
                $py = $set.$node.Item2
                $combo += $hashedBoard[$px][$py]
                $comboCoords.add($i, @($px,$py))
            }

            if ($combo -imatch "]" -or $combo.length -lt 3) {
                continue
            }

            try {
                $boardCombos.Add($combo, @{0 = $comboCoords})
            } catch {
                $boardCombos.$combo.Add($boardCombos.$combo.Count, $comboCoords)
            }
        }

        return $boardCombos
    }

    # Main is main
    Function Main-FindWords () {
        # Find each possible combo
        $pMatches = @()
        foreach ($combo in $boardCombos.Keys) {
            if ($combo.length -le $word.length -and $combo -eq $word.Substring(0,$combo.length)) {
                $pMatches += $combo
            }
        }

        foreach ($combo in $pMatches) {
            foreach ($k in $boardCombos.$combo.Keys) {
                $startTestWord = $combo
                $startWordPos = $startTestWord.Length
                $startCoordSet = $boardCombos.$combo.$k
                $startCheck = Check-NextPos -passedCoords $startCoordSet -word $word -wordPos $startWordPos -testWord $startTestWord

                if ($startCheck) {
                    $script:words.Add($word, $startCheck)
                    return
                }
            }
        }
    }

    Init-FindWords
    Main-FindWords
    
    return "Word not found"
}
#endregion

Init

$GameSelect_Combo.add_SelectionChanged({
    switch($this.selectedvalue){
        "Classic" {$Script:BoardSize = 4}
        "New" {$Script:BoardSize = 4}
        "Big_Original" {$Script:BoardSize = 5}
        "Big_Challenge" {$Script:BoardSize = 5}
        "Big_Deluxe" {$Script:BoardSize = 5}
        "Big_2012" {$Script:BoardSize = 5}
        "Super_Big" {$Script:BoardSize = 6}
    }
})

$StartButton.add_click({
    $script:BoggleBoard = New-BoggleGrid -GameSelection $GameSelect_Combo.SelectedValue
    if ($GameSelect_Combo.SelectedValue -eq "Debugging_Mode") {
        $script:Seconds = $script:DebugSeconds
    }
    $script:boggleJob = Start-Job -ScriptBlock $FindWords -ArgumentList @($BoggleBoard, $WordFile, ($script:Seconds - 10), $BoardSize) -Name "Boggle"

    $Grid.ColumnDefinitions.Clear()
    $Grid.RowDefinitions.Clear()
    0..($BoardSize-1) | %{$Grid.RowDefinitions.Add([System.Windows.Controls.RowDefinition]@{Height="110"})}
    0..($BoardSize-1) | %{$Grid.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]@{Width="110"})}

    $Grid.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]@{Width=40})
    $script:ProgressBar = [System.Windows.Controls.ProgressBar]@{Minimum=1; Maximum=$script:Seconds;Value=$script:Seconds;Foreground="RoyalBlue";Orientation= 1}
    Add-ToGrid -Control $ProgressBar -Parent $Grid -X $BoardSize -Y 0 -RowSpan $BoardSize

    foreach($x in (0..($BoardSize-1))){
        foreach($y in (0..($BoardSize-1))){
            $Border = [System.Windows.Controls.Border]@{BorderBrush="Black";BorderThickness=1;}
            $Border.Name = "c$x$y"
            $Border.AddChild(([System.Windows.Controls.TextBlock]@{Text="";FontSize=80;HorizontalAlignment="Center";VerticalAlignment="Center";Foreground="Black"}))
            Add-ToGrid -Control $Border -Parent $Grid -X $y -Y $x -ZIndex 0
        }
    }

    $This.Visibility = "Collapsed"
    $GameSelect_Combo.Visibility = "Collapsed"
    
    $timer.Add_Tick({
        #Countdown before the start of the game
        if ($countdown -gt 0) {
            $Grid.children | where {$_.name -match "C[0-9][0-9]"} | foreach {$_.child.text = ""}
            foreach ($key in $countdownWords[$countdown].keys) {
                $Grid.Children | where {$_.name -eq $key} | foreach {$_.child.text = $script:countdownWords[$countdown].$key}
            }
            $script:countdown--
        } else {
            # Reset the colors once the countdown is done
            if ($countdown -eq 0) {
                foreach($x in (0..($BoardSize-1))){
                    foreach($y in (0..($BoardSize-1))){
                        $Grid.children | where {$_.name -eq "C$x$y"} | foreach {$_.child.text = $script:boggleBoard[$x][$y]}
                    }
                }
                $script:countdown = $null
            }

            # Incriment the timer bar
            $script:ProgressBar.Value = $script:Seconds--
            [System.Windows.Forms.Application]::DoEvents()

            # Once the findwords job is completed collect the job and order the list alphabetically
            if($boggleJob.State -eq "Completed" -and $script:boggleJob.HasMoreData -eq $true){
                Print-Board $BoggleBoard
                $wordSearchResults = $boggleJob | Receive-Job

                $finalFoundWords = $wordSearchResults[0]
                $script:boardCombos = $wordSearchResults[1]
                $finalFoundWords.Add("[clear]", $null)
                $finalFoundWords.GetEnumerator() | Sort-Object -Property Name | ForEach-Object {$Words.Add($_.Name, $_.Value)}
            }

            if($script:Seconds -le 0){
                # Find the top 5 words and post them
                foreach ($word in $script:Words.keys) {
                    if ($word -eq "[clear]") {continue}
                    Check-WordLength -word $word
                }
                Write-Host ("The top 5 words ScriptBob found were:")
                for ($i=0; $i -lt 5; $i++) {
                    Write-Host ("#" + ($i+1) + "`: " + $script:topFive[$i] + " ("+ $script:topFive[$i].length +")")
                }

                $Grid.Background = "Black"
                $Grid.Children | %{if($_.GetType().Name -eq "Border"){$_.Background = "Black";$_.Child.Foreground = "White";$_.Child.Background = "Black"}}
                $this.Stop()
                $window.width = 1500
                $vb.width = 1500

                $grid.ColumnDefinitions[$Script:BoardSize].width = 240
                $foundWordDisplay = [System.Windows.Controls.Border]@{BorderBrush="Black";BorderThickness=1}
                $foundWordDisplay.Name = "foundWords"

                $script:ListBox = [System.Windows.Controls.ListBox]@{ItemsSource=$script:Words.keys;Height=660;Width=240;FontSize=50;Foreground = "White";Background = "Black"}
                $ListBox.add_SelectionChanged({
                    Highlight-Words -word $this.selectedvalue
                    if ($this.selectedvalue -ne "[clear]") {
                        Update-Definition -word $this.selectedvalue
                    }
                })
                $foundWordDisplay.AddChild($ListBox)
                Add-ToGrid -Control $foundWordDisplay -Parent $Grid -X $BoardSize -Y 0 -RowSpan $BoardSize

                $script:searchBox = [System.Windows.Controls.TextBox]@{Width=300;FontSize=24}
                $script:searchButton = [System.Windows.Controls.Button]@{Width=200;FontSize=24;Content="Search";Background="LightGreen"}
                $script:searchResults = [System.Windows.Documents.FlowDocument]@{}
                $script:searchText = [System.Windows.Documents.Paragraph]@{FontSize=16; Foreground="White"}
                $script:scrollViewer = [System.Windows.Controls.FlowDocumentScrollViewer]@{}
                $searchResults.Blocks.add($searchText)
                $scrollViewer.AddChild($searchResults)
                $Grid.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]@{Width="300"})
                $Grid.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]@{Width="200"})
                
                $searchButton.add_click({
                    Word-Search
                })

                $searchBox.Add_keyDown({
                    param ($sender, $e)

                    if ($e.Key -eq "Return") {
                        Word-Search
                    }
                })
                
                Add-ToGrid -Control $searchBox -Parent $Grid -X ($BoardSize+1) -Y 0
                Add-ToGrid -Control $searchButton -Parent $Grid -X ($BoardSize+2) -Y 0
                Add-ToGrid -Control $scrollViewer -Parent $Grid -X ($BoardSize+1) -Y 1 -RowSpan ($BoardSize-1) -ColumnSpan 2
            }
        }
    })
    $timer.Start()
})

[void]$Window.ShowDialog()

if (Get-Job) {
    Get-Job -Name "Boggle" | Stop-Job
    Get-Job -Name "Boggle" | Remove-Job
}
