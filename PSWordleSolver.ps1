$url = 'https://www.nytimes.com/games/wordle/index.html'

#change to false if command prompt and browser window are needed
$runsilent = $false

# selenium working directory
$workingPath = 'C:\selenium'

# Add the working directory to the environment path.
# This is required for the ChromeDriver to work.
if (($env:Path -split ';') -notcontains $workingPath) {
    $env:Path += ";$workingPath"
}

#load the selenium module
Import-Module "$($workingPath)\WebDriver.dll"

# Create a new ChromeDriver Object instance.
$options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$defaultservice = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService()

#if runsilent is true, the options will be set so that the chrome window and command prompt will not show
if ($runsilent -eq $true){
#change from true to false if you want to see the webdriver command prompt
    $defaultservice.HideCommandPromptWindow = $true
    $options.AddArguments('ignore-certificate-errors','headless')
} else {
    $defaultservice.HideCommandPromptWindow = $false    
    $options.AddArguments('ignore-certificate-errors')
}

#check to see if the chrome driver is already running.  if not, start a new one
if ($ChromeDriver.CurrentWindowHandle -eq $null){
    $ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($defaultservice,$options) #start the chrome driver with options selected above
    
}

#go to web page
$ChromeDriver.Navigate().GoToURL($url)
start-sleep -Seconds 2
$sends = $ChromeDriver.FindElementByXPath('/html/body')
#click inside body of html to close instructions
$sends.click()

##GLOBALS
[System.Collections.ArrayList]$words = Get-Content 'wordlist.txt'
$perfect = @{}
$bad = @{}
$good = @{}
$inp = $null
$ing = $null
$inpv = $null
$ingv = $null
$inb = $null
$inv = $null
$row1 = $null
$row2 = $null
$row3 = $null
$row4 = $null
$row5 = $null
$row6 = $null
$row = @()
$i = $null
$count = 1

#sends the new word
function sendkeys{
    param(
        $word
    )

    $sends.SendKeys($word)
    $sends.SendKeys([OpenQA.Selenium.Keys]::Enter)

    getrowdata
}

#captures the latest row data and sends it for evaluation
function getrowdata{
    
    $row1 = $ChromeDriver.ExecuteScript("return document.querySelector('game-app').shadowRoot.querySelector('game-row:nth-child(1)').shadowRoot.querySelectorAll('game-tile[letter]')") |foreach {$_.GetAttribute("outerHTML")}
    $row2 = $ChromeDriver.ExecuteScript("return document.querySelector('game-app').shadowRoot.querySelector('game-row:nth-child(2)').shadowRoot.querySelectorAll('game-tile[letter]')") |foreach {$_.GetAttribute("outerHTML")}
    $row3 = $ChromeDriver.ExecuteScript("return document.querySelector('game-app').shadowRoot.querySelector('game-row:nth-child(3)').shadowRoot.querySelectorAll('game-tile[letter]')") |foreach {$_.GetAttribute("outerHTML")}
    $row4 = $ChromeDriver.ExecuteScript("return document.querySelector('game-app').shadowRoot.querySelector('game-row:nth-child(4)').shadowRoot.querySelectorAll('game-tile[letter]')") |foreach {$_.GetAttribute("outerHTML")}
    $row5 = $ChromeDriver.ExecuteScript("return document.querySelector('game-app').shadowRoot.querySelector('game-row:nth-child(5)').shadowRoot.querySelectorAll('game-tile[letter]')") |foreach {$_.GetAttribute("outerHTML")}
    $row6 = $ChromeDriver.ExecuteScript("return document.querySelector('game-app').shadowRoot.querySelector('game-row:nth-child(6)').shadowRoot.querySelectorAll('game-tile[letter]')") |foreach {$_.GetAttribute("outerHTML")}

    $count = 1

    if($row6){
        #send row
        #write-host "sending row 6"
        parse -row $row6
    } else {
        if($row5){
            #send row
            #write-host "sending row 5"
            parse -row $row5
        } else {
            if($row4){
                #send row
                #write-host "sending row 4"
                parse -row $row4
            } else {
                if($row3){
                    #send row
                    #write-host "sending row 3"
                    parse -row $row3
                } else {
                    if($row2){
                        #send row
                        #write-host "sending row 2"
                        parse -row $row2
                    } else {
                        if($row1){
                            #send row
                            #write-host "sending row 1"
                            parse -row $row1
                        }
                    }
                }
            }
        }
    }
}

#evaluates based on present,absent,or correct, and adjusts the word list
function parse{
    param(
        $row = @()
    )

    foreach($l in $row){

        if(($l -split '"')[7] -eq 'win'){
            write-host "you win!"
        } else {
            $position = $count
            $letter = ($l -split '"')[1]
            $evaluation = ($l -split '"')[3]
                        
            if($evaluation -eq 'present'){
                #update good letters
                update-letters -g $letter -gv $position
                #write-host updated good letters
                $count++
            } else {
                
                if($evaluation -eq 'correct'){
                    #update perfect letters
                    update-letters -p $letter -pv $position
                    #write-host updated perfect letters
                    $count++
                } else {
                    
                    #update bad letters
                    update-letters -b $letter -bv $position
                    #write-host updated bad letters
                    $count++
                
                }
            }
        }
    }
    evaluate
}

function update-letters{
    #function used to update good, bad, and perfect letters with letter positions
    param(
        $p, #perfect letter
        [string]$b, #letter
        $g, #good letter
        $pv, #perfect position
        $gv, #good position
        $bv #bad position
    )

    if($p){$perfect.$p = $pv}
    if($g){$good.$g = $gv}
    if($b){$bad.$b = $bv}
}

function evaluate{
    
    #remove words that dont include perfect letters
    foreach($word in $($words.clone())){
        foreach($h in $perfect.GetEnumerator()){
            [int]$position = $h.value
            [char]$letter = $h.name

            if($word[$position - 1] -eq $letter){
                #word can stay
                #write-host $letter "matches" $word[$position - 1]
                
            } else {
                #word needs deleted
                #write-host $word "removed"
                $words.remove($word)
            }
        }

    }
    #END FOREACH

    #remove bad words
    foreach($word in $($words.clone())){
        foreach($h in $bad.GetEnumerator()){
            [int]$position = $h.value
            [string]$letter = $h.name

            #the below code removes "bad" letters based on position rather than outright.  it's safer this way
            if($word[$position -1] -eq $letter){
                #word needs deleted
                #write-host $word "removed"
                $words.remove($word)
             }



             #the below code removes any letter that comes back absent but causes issues from time to time 
             <#
             if($word.IndexOfAny($letter) -ge 0){
                #word needs deleted
                $words.remove($word)
                
            }
            #>
        }
            
    }
    #END FOREACH

    #remove words where the position of good letters has already been attempted
    foreach($word in $($words.clone())){
        foreach($h in $good.GetEnumerator()){
            [int]$position = $h.value
            [char]$letter = $h.name

            #if the word has a known good letter in the wrong spot, remove it
            if($word[$position - 1] -eq $letter){
                #word needs deleted
                #write-host $word "removed"
                $words.remove($word)
                
            }

            #if the word does not contain a known good letter, remove it
            if($word.IndexOfAny($letter) -eq -1){
                #word needs deleted
                #write-host $letter "matches" $word[$position - 1]
                #write-host $word "removed"
                $words.remove($word)
            }
        }

    }

}

function guess {
    #return list of possible words
    $guess = $words |Get-Random
    sendkeys -word $guess

    return write-host ($words|measure).count "words left"
}

start-sleep -seconds 1

$q = 0
while ($q -lt 6){
    start-sleep -seconds 2 
    guess
    $q++
    
}

function reset{
    $ChromeDriver.quit()

}
