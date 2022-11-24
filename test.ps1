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
        $NumOfSentences = ($Text.ToCharArray() | Where-Object {$_ -eq '.'} | Measure-Object).Count # Przejscie na tablice char i zliczenie kropek
        $NumOfAllSentences+=$NumOfSentences # Zwiekszenie odpowiednio wszystkich zliczonych kropek
        $TextOnlyWords = $Text -Replace"[^\x20\x41-\x5A\x61-\x7A]+", "" # Pozbycie się wszystkich znakow oprocz " ", "A"-"Z", "a"-"z".
        # https://www.asciitable.com
        $Words=$TextOnlyWords.split(" ")
        $NumOfWords=$Words.Count
        $NumOfAllWords+=$NumOfWords

        
        #foreach($Sentence in $Sentences)
        #{
               # Write-Output $Sentence 
        #}

    }
    Write-Output $NumOfAllWords
    Write-Output $NumOfAllSentences
    $Document.close()
    $Word.Quit()


}
Get-Tree -Path D:\testy\wwww.docx  # Uruchomienie funckji 