# **Technologie i Jakość Oprogramowania**

## Autor: Marcel Nosek

## Temat: Szybka opinia

### Opis:

Szybka Opinia to aplikacja wspierająca proces tworzenia opinii technicznych w sposób szybki i niemal w pełni automatyczny.
Użytkownik podaje podstawowe informacje, takie jak dane zlecającego, odbiorcy oraz identyfikator sprzętu, a następnie wybiera typ opinii.
Na tej podstawie program automatycznie generuje niezbędne pliki, w tym dokumenty Word oraz arkusze Excel, zgodnie z wybranym wariantem opinii.
Celem projektu jest ograniczenie ręcznej pracy, przyspieszenie przygotowania dokumentacji oraz zachowanie spójnych standardów tworzonych opinii.

## **Uruchomienie**

Program można uruchomić wpisując komendę *python main.py* bedąc w katalogu głównym lub korzystając z pliku main.exe znajdującym się w katalogu głównym.

## **API**

### Główne działanie programu
Program wyszukuje dane o sprzęcie na podstawie podanego klucza (np. ID). Niezależnie od ilości znalezionych danych przygotowuje szablon z polami do wpisywania danych (dla każdego sprzętu generuje się osobny szablon). Jeśli znalazł wczesniej dane to pola są odpowiednio uzupełniane. Użytkownik może edytować dane. Gdy użtywkonik kliknie przycisk "Generuj opinię"
program automatycznie tworzy plik word z odpowiednimi formułami dotyczącymi sprzętu zarówno dla likwidacji jak i wyceny. W przypadku wycen tworzone są dodatkowo
kosztorysy. Użytkowik wybierając opcje opinii z kosztorysem jest w stanie uzupełnić m.in. linki i ceny sprzętów o podobnych parametrach, aby tego dane pojawiły się przy generacji kosztorysu.

### Przeszukiwane pliki:
Poniższe pliki excel są przesukiwane podczas szukania informacji o sprzęcie na podstawie podanego ID.
Jeśli program nie znajdzi urządzenia lub danych z nim związanych może to oznaczać brak tych danych w żadnym z podanym plików:

- "Ewidencja.xlsx"
- "Log1.xlsx"
- "Log2.xlsx"
- "Log3.xlsx"

Dodatkowo program przeszukując pliki sprawdza tylko aktywny kosztorys wiec należy się upewnić że nie ma w nim zbędnych arkuszy lub czy odpowiedni arkusz jest zapisany jako ostatnio aktywny.
Program wciąż przygotuje formularz opinii pomimo braku danych wieć nie jest to niezbędne do jego działania, jednak może się to wiązać z brakiem istotnych danych.

### Kolumny danych

Każdy z arkuszy powinien posiadać oprócz kolumn z danymi przynajmniej jedną kolumnę z dostepnych kluczy. Brak klucza spowoduje że porgram nie będzie miał opcji znaleść danych w tym pliku.

Kolumny kluczy:

- "Numer ewidencyjny" - ID sprzętu
- "Numer środka trwałego" - SAP sprzętu
- "Nr inwentarzowy" - numer inwentarzowy sprzętu

Kolumny danych

- "Nazwa" - nazwa użytkownika sprzętu
- "Typ" - rodzaj sprzętu (np. komputer)
- "Producent" - producent sprzętu
- "Model" - model sprzętu
- "Numer seryjny" - numer seryjny sprzętu
- "Data zakupu" - data zakupu sprzętu
- "System operacyjny" - SO sprzętu
- "Model procesora" - CPU sprzętu
- "Pamięć HDD (GB)" - Łączna pamięć dyskowa sprzętu
- "Pamięć RAM (GB)" - Łączny RAM sprzętu

### Plik Departaments
---
Program ma zaimplementowaną listę która jest uzupełniania działami z pliku ***"Departments.xlsx"*** Plik ten zawiera zbiór działów Grupa Azoty S.A. Dzięki niemu jesteśmy w stanie kontorlować listę dostępnych działów przy tworzeniu opinii.

Wymagane kolumny:
- "Kod działu"
- "Nazwa działu"

### Opis:	
W arkuszu dane są zapisane w układzie hierarchicznym, kolumnami schodkowo. Każdy kolejny poziom hierarchii (poddział) znajduje się w następnej parze kolumn (np. dział główny w A-B, jego poddział w C-D, poddział tego poddziału w E-F itd.). Dzięki temu struktura tworzy przejrzyste drzewo organizacyjne. Jeśli chcesz dodać nowy poddział, umieść go w odpowiedniej 
kolumnie zgodnie z poziomem hierarchii, tak aby zachować przejrzystość i spójność całej struktury. Kolumny z kodami działów zostały nazwane, jako "Kod działu", a obok znajduję się kolumna, "Nazwa działu" z ich pełnymi nazwami.

Przykładowa hierarchia:
```
kod - nazwa - kod - nazwa - kod - nazwa
              kod - nazwa
              kod - nazwa - kod - nazwa
              kod - nazwa
kod - nazwa - kod - nazwa - kod - nazwa
              kod - nazwa - kod - nazwa
              kod - nazwa
```

### Inne pliki i foldery
- **Other_files** - Folder zawierający dodatkowe pliki które są wykorzystywane przy generacji opinii 
- **logo.png** - Logo które zostanie umieszone w opinii
- **Cennik.pdf** - Przygotowany zrzut w formie pdf. Zawiera stronę z cenami wymiany chłodzenia( potrzebne do wycen sprzętów )
- **Wygenerowane_opinie** - foder zawierjący wygenerowane opinie przez program Szybka Opinia ( program sam go tworzy jeśli go brak )
- **Tests** - Folder zawierający testy jednostkowe oraz integracyjne projektu
- **Bazy_danych** - Zawiera pliki któe są przeszukiwane w celu wydobycia danych
- **main.exe** - plik wykonywalny aplikacji

## **Testy**

### Uruchamianie testów
Wszystkie testy możemy uruchomić korzystając z komendy:
> python -m unittest discover -s Tests

Jeśli chcemy uruchomić konkretny plik testowy korzystamy z :
> python -m unittest Tests.{nazwa pliku testu}

Jeśli chcemy uruchomić konretny pojedynczy test korzystamy z:
> python -m unittest Tests.{nazwa pliku testu}.{nazwa klasy testu}.{nazwa testu}

### Testy jednostkowe
1. test_getDecision_liquidation - test sprawdza czy przy wyborze likwidacji wybór decyzji zwróci false (oznacza to żeby nie robić korztorysu)
> Lokalizacja: Tests/test_device.py

2. test_getDecision_valudation - test sprawdza czy przy wyborze wyceny wybór decyzji zwróci true (oznacza to żeby przygotowac korztorys)
> Lokalizacja: Tests/test_device.py

3. test_validateRequiredFields_missing - test sprawdza czy przy braku wpisanego klucza wyskakuje komunikat o jego braku i czy wiadomość pojawia się tylko raz.
> Lokalizacja: Tests/test_device.py

4. test_updateDeviceInfo_updates_info - test sprawdza czy wpisanie recznie dane zostaną odpowiednio przekazane do device_info (potrzebne przy odczycie danych do plików)
> Lokalizacja: Tests/test_device.py

5. test_format_reporter_name_standard - test sprawdza popranwość formatowania nazwy osoby proszącej o opinię gdy podano dwa słowa (imię i nazwisko).
> Lokalizacja: Tests/test_form.py

6. test_format_reporter_name_many_parts - test sprawdza popranwość formatowania nazwy osoby proszącej o opinię gdy podano nazwę składającą się z więcej niż 2 słów.
> Lokalizacja: Tests/test_form.py

7. test_format_reporter_name_single_word - test sprawdza popranwość formatowania nazwy osoby proszącej o opinię gdy podano tylko jedno słowo.
> Lokalizacja: Tests/test_form.py

8. test_format_author_name - test sprawdza popranwość formatowania nazwy autora opinii gdy podano dwa słowa (imię i nazwisko).
> Lokalizacja: Tests/test_form.py

9. test_format_author_name_single - test sprawdza popranwość formatowania nazwy autora opinii gdy podano jedno słowo.
> Lokalizacja: Tests/test_form.py

10. test_create_opinion_fails_on_missing_device_id - test sprawdza czy id urządzena pozostało wpisane (użytkownik mógł je usunąć).
> Lokalizacja: Tests/test_request.py

11. test_create_opinion_fails_on_invalid_required_fields - test sprawdza czy po braku wymaganych pól program odpowiednio się zachowuje (nie uruchamia kolejnych funkji). Test potwierdza wcześniejsze wyłapanie błedu.
> Lokalizacja: Tests/test_request.py

12. test_create_opinion_success - test sprawdza czy opinia została poprawnie stworzona.
> Lokalizacja: Tests/test_request.py

13. test_create_opinion_exception_sets_opinion_none - test sprawdza zachowanie programu w przypadkku wystąpienia błedu przy tworzeniu opinii.
> Lokalizacja: Tests/test_request.py

14. test_create_pricing_called_when_decision_true - test sprawdza czy przy wyborze stworzenia kosztorysu jest on faktycznie tworzony.
> Lokalizacja: Tests/test_request.py

15. test_create_pricing_not_called_when_decision_false - test sprawdza czy przy nie wybraniu tworzenie kosztorysu faktycznie nie zostanie on stworzony.
> Lokalizacja: Tests/test_request.py

16. test_create_pricing_exception_is_handled - test sprawdza czy w przypadku błedu program zachowuje się prawidłowo. Wyskakuje komunikat o błędzie i program nie crashuje.
> Lokalizacja: Tests/test_request.py

### Testy Integracyjne

1. test_findDataByKey_fills_info - test sprawdza czy program jest w stanie szukając po podanym kluczu znaleść dane dla urządzenia w excelu
> Lokalizacja: Tests/test_device.py

2. test_findData_shows_warning_on_missing_fields - test sprawdza czy wyskakuje komunikat o braku danych o szukanym urządzeniu. 
> Lokalizacja: Tests/test_device.py

3. test_get_departments_mocked - test ma na celu sprawdzić czy lista departamentów jest odpowiednio pobierana i przygotowana w postaci listy.
> Lokalizacja: Tests/test_form.py

4. test_createOpinion_stops_on_first_invalid_device - test sprawdza czy tworzeniu opini zatrzymuje się przy pierwszym błędnym urządzeniu (urządzeniu o błednych lub niepełnych danych).
> Lokalizacja: Tests/test_request.py

5. test_updateDeviceInfo_called_for_all_devices - test sprawdza czy wszystkie urządzenia aktualizuja swoje dane przed przygotowaniem opinii
> Lokalizacja: Tests/test_request.py

6. test_pricing_created_only_for_devices_with_decision - test sprawdza czy kosztorys jest przygotowywany tylko dla urządzeń przy których został wybrany.
> Lokalizacja: Tests/test_request.py

7. test_pricing_exception_does_not_stop_other_pricing - test sprawdza czy gdy pojawi się problem przy tworzeniu jednego kosztorysu pozostałe i tak zostaną stworzone.
> Lokalizacja: Tests/test_request.py

8. test_creates_docx_file - test sprawdza plik opinii został stworzony w wyznaczonym miejscu.
> Lokalizacja: Tests/test_opinion.py

9. test_pricing_creates_correct_directory - test sprawdza czy folder kosztorysu jest tworzony w odpiwednim miejscu.
> Lokalizacja: Tests/test_pricing.py

10. test_pricing_copies_cennik_pdf - test sprawdza czy w folderze kosztorysu tworzony jest plik Cennik.pdf wymagany przy każdym kosztorysie.
> Lokalizacja: Tests/test_pricing.py

11. test_pricing_creates_xlsx_file - test sprawdza czy w folderze kosztorysu tworzony jest plik kosztorysu (rozszerzenie xlsx).
> Lokalizacja: Tests/test_pricing.py

## Scenariusze testowe (TestCase)

### T01 – Test wyświetlania błędu o braku wpisanego ID


**Warunki początkowe:**  
Aplikacja jest uruchomiona

**Kroki testowe:**
1. Spróbuj użyć przycisku „Generuj”, aby przejść dalej.

**Oczekiwany rezultat:**  
Bez wpisanego ID wyskakuje okienko z błędem o jego braku.  
Nie można wygenerować formularza z danymi bez uzupełnienia ID.

---

### T02 – Test komunikatu o braku części danych w formularzu startowym


**Warunki początkowe:**  
Aplikacja jest uruchomiona

**Kroki testowe:**
1. Uzupełnij tylko pole „ID”.
2. Spróbuj użyć przycisku „Generuj”, aby przejść dalej.
3. Jeśli wyskoczy komunikat, kliknij przycisk „Tak”.

**Oczekiwany rezultat:**  
Przy niewypełnieniu części danych wyskakuje komunikat z pytaniem, czy na pewno chcemy wygenerować dalszą część bez uzupełnionych danych.  
Po potwierdzeniu przechodzimy bez problemu do okna wygenerowanego formularza.

---

### T03 – Test znajdowania danych na podstawie wpisanego ID

**Warunki początkowe:**  
Aplikacja jest uruchomiona

**Kroki testowe:**
1. Sprawdź czy przygotowany jest plik z danymi lub umieść plik Excel z danymi w folderze `Bazy_danych`, zmieniając jego nazwę na `Log1.xlsx`. Dane powinny być zgodne ze standardem opisanym w dokumentacji.
2. Wpisz w polu „ID” jeden z numerów znajdujących się w pliku z danymi.
3. Kliknij przycisk „Generuj”.
4. Sprawdź wygenerowany formularz.

**Oczekiwany rezultat:**  
Wygenerowany formularz zawiera dane adekwatne do wpisanego ID.  
Zawartość formularza zgadza się z danymi z pliku Excel.

---

### T04 – Test listy departamentów

**Warunki początkowe:**  
Aplikacja jest uruchomiona

**Kroki testowe:**
1. Sprawdź czy przygotowany jest plik z działami lub umieść plik Excel z działami w folderze `Bazy_danych`, zmieniając jego nazwę na `Departments.xlsx`. Dane powinny być zgodne ze standardem opisanym w dokumentacji.
2. W aplikacji rozwiń pole „Dział” i sprawdź, czy lista zgadza się z danymi w pliku Excel.
3. Wpisz kod jednego z działów lub fragment jego nazwy i kliknij Enter.
4. Sprawdź, jakie działy zostały wyświetlone.

**Oczekiwany rezultat:**  
Lista zawiera wszystkie działy wraz z ich kodami zgodnie z zawartością pliku Excel.  
Po wpisaniu części nazwy lub kodu wyświetlane są tylko pasujące elementy.

---

### T05 – Test przemieszczania się między oknami aplikacji

**Warunki początkowe:**  
Aplikacja jest uruchomiona, przygotowane są pliki z danymi i działami

**Kroki testowe:**
1. Wpisz dowolne ID i kliknij przycisk „Generuj”.
2. Po przejściu do nowego widoku kliknij przycisk „Wstecz”.
3. Po powrocie do głównego widoku sprawdź, czy przycisk „Dalej” się aktywował.
4. Kliknij przycisk „Dalej”.

**Oczekiwany rezultat:**  
Przycisk „Dalej” jest początkowo nieaktywny.  
Po wygenerowaniu formularza można wrócić do poprzedniego widoku przyciskiem „Wstecz”.  
Przycisk „Dalej” staje się aktywny i przenosi użytkownika do wygenerowanego formularza bez utraty danych.

---

### T06 – Test dodawania i usuwania urządzeń

**Warunki początkowe:**  
Aplikacja jest uruchomiona

**Kroki testowe:**
1. Kliknij przycisk „Dodaj urządzenie”.
2. Kliknij czerwony przycisk „X” przy dowolnym urządzeniu.

**Oczekiwany rezultat:**  
Urządzenia są natychmiast dodawane lub usuwane z formularza zgodnie z wykonanymi akcjami.

---

### T07 – Test generowania standardowej opinii

**Warunki początkowe:**  
Aplikacja jest uruchomiona, przygotowane są pliki z danymi i działami

**Kroki testowe:**
1. Uzupełnij wszystkie pola formularza głównego.
2. Kliknij przycisk „Generuj”.
3. W wygenerowanym formularzu kliknij przycisk „Generuj opinię”.

**Oczekiwany rezultat:**  
Wyświetla się komunikat o wygenerowaniu opinii.  
W folderze `Wygenerowane_opinie` znajduje się folder z nową opinią zawierający plik opinii zgodny ze standardami firmy.

---

### T08 – Test zmiany rodzaju opinii

**Warunki początkowe:**  
Aplikacja jest uruchomiona

**Kroki testowe:**
1. Wpisz dowolne ID i kliknij przycisk „Generuj”.
2. Wybierz rodzaj opinii „Wycena (Likwidacja)”.
3. Wybierz rodzaj opinii „Wycena (Przekazanie)”.
4. Wybierz rodzaj opinii „Likwidacja”.

**Oczekiwany rezultat:**  
Zmiana rodzaju opinii powoduje odpowiednie zmiany w formularzu, m.in. w opisie wyceny oraz aktywacji lub dezaktywacji pól.

---

### T09 – Test generowania opinii z wyceną

**Warunki początkowe:**  
Aplikacja jest uruchomiona

**Kroki testowe:**
1. Uzupełnij wszystkie pola formularza głównego.
2. Kliknij przycisk „Generuj”.
3. Zmień rodzaj opinii na wycenę.
4. Uzupełnij pole „Numer wyceny”.
5. Kliknij przycisk „Generuj opinię”.

**Oczekiwany rezultat:**  
Dla wygenerowanej opinii tworzony jest dodatkowy folder z wyceną.  
W folderze znajdują się pliki Excel oraz PDF z cennikiem zgodne ze standardami firmy.

---

### T10 – Test generowania opinii z wyceną oraz zmianami użytkownika

**Warunki początkowe:**  
Aplikacja jest uruchomiona

**Kroki testowe:**
1. Uzupełnij wszystkie pola formularza głównego.
2. Kliknij przycisk „Generuj”.
3. Zmień rodzaj opinii na wycenę.
4. Uzupełnij pole „Numer wyceny”.
5. Wprowadź zmiany w formularzu lub uzupełnij wcześniej puste pola.
6. Kliknij przycisk „Generuj opinię”.
7. Sprawdź zawartość wygenerowanych plików.

**Oczekiwany rezultat:**  
Pliki opinii oraz wyceny są generowane zgodnie ze standardami firmy.  
Wszystkie zmiany wprowadzone przez użytkownika są poprawnie odwzorowane w wygenerowanych plikach.

## **Użyte Technologie**

### Główna technologia

**Python** - język programowania wykorzystany do stworzenia całej aplikacji, odpowiedzialny za logikę, obsługę danych oraz interfejs użytkownika.

### Biblioteki

**datetime** - biblioteka standardowa służąca do operacji na datach i czasie, wykorzystywana m.in. do zapisu znaczników czasowych i formatowania dat.

**docx** - biblioteka umożliwiająca tworzenie oraz edycję dokumentów Microsoft Word (.docx) bezpośrednio z poziomu aplikacji.

**openpyxl** - narzędzie do odczytu, zapisu i modyfikacji plików Excel (.xlsx), wykorzystywane do pracy z danymi tabelarycznymi.

**os** - biblioteka standardowa zapewniająca dostęp do funkcji systemu operacyjnego, takich jak operacje na plikach, katalogach oraz zmiennych środowiskowych.

**re** - biblioteka do obsługi wyrażeń regularnych, używana do wyszukiwania, walidacji i przetwarzania tekstu.

**shutil** - biblioteka standardowa umożliwiająca zaawansowane operacje na plikach i katalogach, takie jak kopiowanie, przenoszenie i usuwanie danych.

**tkinter** - standardowa biblioteka Pythona służąca do tworzenia graficznego interfejsu użytkownika (GUI), odpowiedzialna za okna, przyciski i formularze aplikacji.