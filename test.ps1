function Get-Tree 
{
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)] # Parametr obowiązkowy to ścieżka
        [string] $Path # Ustawianie typu parametru na string
    )
    $Word = New-Object -ComObject Word.application  # Przypisanie obiektu worda do zmiennej
    $Word.Visible = $False # Ustawienie by otwarcie nie bylo widoczne
    $Document = $word.Documents.Open($Path) # Otworzenie dokumentu
    $Paragraphs = $Document.Paragraphs # Pobranie danych wszystkie paragrafy
    $NumOfAllSentences = 0 # Liczba wszystkich zdan
    $NumOfAllWords = 0 # Liczba wszystkich slow 
    $NumOfWordsPerParTab = @() # Tablica z iloscia slow w danym paragrafie
    $LongestWord = " " # Sluzy do znalezienia najdluzszego slowa
    $LongestSentence = @() # Sluzy do znalezienia najdluzszego zdania
    foreach ($Paragraph in $Paragraphs) 
    {
        $Text = $Paragraph.range.Text # Tekst z każdego paragrafu
        $Text = $Text.Replace("...", "") # Usnięcie "..." w celu zliczenia ".", ktore oznaczaja ilosc zdan
        $Text = $Text.Replace(",", "") # Usnięcie "," poniewaz nie sa do niczego potrzebne a tak zaliczaly by sie do slow
        $NumOfSentences = ($Text.ToCharArray() | Where-Object { $_ -eq '.' -or $_ -eq '?' } | Measure-Object).Count # Przejscie na tablice char i zliczenie kropek lub ?
        $NumOfAllSentences += $NumOfSentences # Zwiekszenie odpowiednio wszystkich zliczonych kropek
        $Sentences = $Text.split("?.")
        foreach ($Sentence in $Sentences) 
        {
            $WordInSen = $Sentence.split(" ")
            if ($WordInSen.Count -gt $LongestSentence.Count) 
            {
                $LongestSentence = $WordInSen  
            }
        }
        $TextOnlyWords = $Text -Replace "[^\x20\x41-\x5A\x61-\x7A]+", "" # Pozbycie się wszystkich znakow oprocz " ", "A"-"Z", "a"-"z". https://www.asciitable.com
        $Words = $TextOnlyWords.split(" ") # Podzielnie paragrafu na pojedyncze wyrazy
        $NumOfWords = $Words.Count # Zliczenie wyrazow w paragrafie
        $NumOfWordsPerParTab += $NumOfWords # Zapisanie ilosci slow w tablicy
        $NumOfAllWords += $NumOfWords # Dodanie do sumy wszystkich wyrazow w dokumencie
        $NewLongestWord = $Words | sort length -desc | select -first 1 # Znalezienie najdluzszego slowa w danym paragrafie 
        if ($NewLongestWord.Length -gt $LongestWord.Length)
        {
            # Sprawdzenie czy jest dluzsze    
            $LongestWord = $NewLongestWord # Przypisanie nowej wartosci slowa
        }
    }
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") # Zaladowanie 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") # Zaladowanie
    $Window = New-Object System.Windows.Forms.Form # Stworzenie nowego okna
    $Window.Text = ("Analyzed " + $Path) # Nazwa okna
    $Window.Size = New-Object System.Drawing.Size(1920 , 1080) # Rozmiary okna
    $Window.StartPosition = "CenterScreen" # Ustawienie pozycji poczatkowej okna
    $Window.TopMost = $true # Ustawienie ze caly czas na wierzchu 
    $Window.Add_Shown({ $Window.Activate() }) # Ustwaienie widocznosci
    $Window.AutoScroll = $True # Wlaczenie opcji scrollowania
    $Graphics = $Window.createGraphics() # Stworzenie grafiki
    $ChartLabel = New-Object system.Windows.Forms.Label # Stworzenie lejbela do wykresu
    $ChartLabel.AutoSize = $true # Ustawienie autosize
    $ChartLabel.Width = 100 # Ustawinie szerokosci
    $ChartLabel.Height = 150 # Ustawienie wysokosci
    $ChartLabel.Font = 'Microsoft Sans Serif,10' # Ustawienie czcionki
    $ChartLabel.Text = ("Ilosc wszystkich slow: " + $NumOfAllWords + "`n") # Tesk do wypisania
    $ChartLabel.Text += ("Ilosc wszystkich zdan: " + $NumOfAllSentences + "`n")
    $ChartLabel.Text += ("Ilosc wszystkich paragrafow: " + $Paragraphs.Count + "`n")
    $ChartLabel.Text += ("Srednia ilosc slow na zdanie: " + ([math]::Round($NumOfAllWords / $NumOfAllSentences)) + "`n")
    $ChartLabel.Text += ("Najdluzszy paragraf ma: " + ($NumOfWordsPerParTab | measure -Maximum).Maximum + " slow" + "`n")
    $ChartLabel.Text += ("Najdluzsze slowo to: '" + $LongestWord + "'" + " i ma ono " + $LongestWord.Length + " liter" + "`n")
    $ChartLabel.Text += ("Najdluzsze zdanie to: '")
    foreach ($WordInLongSen in $LongestSentence) 
    {
        $ChartLabel.Text += ($WordInLongSen + " ")
    }
    $ChartLabel.Text += ("' i ma ono " + $LongestSentence.Length + " wyrazow")
    $Window.Controls.AddRange(@($ChartLabel)) # Wypisanie
    $Window.add_paint(
        {           
            for ($i = 0; $i -lt $NumOfWordsPerParTab.Count; $i++) 
            {
                $ChartLabel = New-Object system.Windows.Forms.Label # Stworzenie lejbela do wykresu
                $ChartLabel.AutoSize = $false # Ustawienie autosize
                $ChartLabel.Width = 150 # Ustawinie szerokosci
                $ChartLabel.Height = 50 # Ustawienie wysokosci
                $ChartLabel.Font = 'Microsoft Sans Serif,10' # Ustawienie czcionki
                $ChartLabel.Text = ("Procent slow w paragrafie" + ($i + 1) + " dla calego tekstu") # Tesk do wypisania
                $Where = $i * 50 + 150 # Obliczenei gdzie ma być wstawiony tekst
                $ChartLabel.Location = New-Object System.Drawing.Point((0, $Where)) # Stworzenie punktu i ustalenie gdzie ma byc wpisany tekst
                $Window.Controls.AddRange(@($ChartLabel)) # Wypisanie
                $Brush = new-object Drawing.SolidBrush green # Stworzenie pedzla
                $Graphics.FillRectangle($Brush, 150, $Where, ($NumOfWordsPerParTab[$i] / 854 * 500), 25) # Naryswonie prostokata
                $Brush = new-object Drawing.SolidBrush red # Stworzenie pedzla
                $Graphics.FillRectangle($Brush, 150 + ($NumOfWordsPerParTab[$i] / 854 * 500), $Where, (500 - ($NumOfWordsPerParTab[$i] / 854 * 500)), 25) # Naryswonie prostokata
                $ChartLabel = New-Object system.Windows.Forms.Label # Stworzenie lejbela do wykresu
                $ChartLabel.AutoSize = $true # Ustawienie autosize
                $ChartLabel.Width = 100 # Ustawinie szerokosci
                $ChartLabel.Height = 50 # Ustawienie wysokosci
                $ChartLabel.Font = 'Microsoft Sans Serif,10' # Ustawienie czcionki
                $ChartLabel.Text = (" " + ($NumOfWordsPerParTab[$i] / 854 * 100) + ' %') # Tesk do wypisania
                $ChartLabel.Location = New-Object System.Drawing.Point((650, $Where)) # Stworzenie punktu i ustalenie gdzie ma byc wpisany tekst
                $Window.Controls.AddRange(@($ChartLabel)) # Wypisanie
            }
        }
    )
    [void] $Window.ShowDialog( ) # Pokazanie okna
    $Document.close() # Zamkniecie doca
    $Word.Quit() # Zamkniecie word
}
Get-Tree -Path D:\testy\wwww.docx  # Uruchomienie funckji 