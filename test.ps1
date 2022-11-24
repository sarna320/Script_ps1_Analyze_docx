function Get-Tree {
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)] # Parametr obowiązkowy to ścieżka
        [string] $Path # Ustawianie typu parametru na string
    )
    $Word = New-Object -ComObject Word.application  # Przypisanie obiektu worda do zmiennej
    $Word.Visible = $False # Ustawienie by otwarcie nie bylo widoczne
    $Document = $word.Documents.Open($Path) # Otworzenie dokumentu

    $Paragraphs = $Document.Paragraphs # Pobranie danych wszystkie paragrafy
    $NumOfAllSentences=0 # Liczba wszystkich zdan
    $NumOfAllWords=0 # Liczba wszystkich slow
    
    foreach ($Paragraph in $Paragraphs)
    {
        $Text = $Paragraph.range.Text # Tekst z każdego paragrafu
        $Text = $Text.Replace("...","") # Usnięcie "..." w celu zliczenia ".", ktore oznaczaja ilosc zdan
        $Text = $Text.Replace(",","") # Usnięcie "," poniewaz nie sa do niczego potrzebne a tak zaliczaly by sie do slow
        $NumOfSentences = ($Text.ToCharArray() | Where-Object {$_ -eq '.' -or $_ -eq '?'} | Measure-Object).Count # Przejscie na tablice char i zliczenie kropek lub ?
        $NumOfAllSentences+=$NumOfSentences # Zwiekszenie odpowiednio wszystkich zliczonych kropek
        $TextOnlyWords = $Text -Replace"[^\x20\x41-\x5A\x61-\x7A]+", "" # Pozbycie się wszystkich znakow oprocz " ", "A"-"Z", "a"-"z".
        # https://www.asciitable.com
        $Words=$TextOnlyWords.split(" ") # Podzielnie paragrafu na pojedyncze wyrazy
        $NumOfWords=$Words.Count # Zliczenie wyrazow w paragrafie
        $NumOfAllWords+=$NumOfWords # Dodanie do sumy wszystkich wyrazow w dokumencie
        $Sentences= $Text.Split('.?') # Podzielenie paragrafu na zdania

        foreach($Sentence in $Sentences)
        {
                $WordsInSent =$Sentence.Split(" ") # Podzielenie zdania na wyrazy
                foreach($WordInSent in $WordsInSent)
                {
                    #Write-Output $WordInSent
                }
                 
        }

    }
    Write-Output $NumOfAllWords
    Write-Output $NumOfAllSentences
    $AverageOfWordsPerSent =[math]::Round($NumOfAllWords/$NumOfAllSentences)
    Write-Output $AverageOfWordsPerSent
    $Document.close()
    $Word.Quit()


}
Get-Tree -Path D:\testy\wwww.docx  # Uruchomienie funckji 