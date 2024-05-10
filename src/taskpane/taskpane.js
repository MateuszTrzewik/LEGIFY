/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    
    if (sideloadMsg) {
      sideloadMsg.style.display = "none";
    }
    
    if (appBody) {
      appBody.style.display = "flex";
    }

    // Referencje do przycisków
    const plButton = document.getElementById("plButton");
    const engButton = document.getElementById("engButton");
    const krsInput = document.getElementById("krsInput");

    // Dodanie nasłuchiwacza na zmianę wartości w polu krsInput
    krsInput.addEventListener('input', () => {
      const krsValue = krsInput.value;
      const isValidKrs = krsValue.length === 10 && /^\d+$/.test(krsValue); // Sprawdzenie czy wartość to dokładnie 10 cyfr
      plButton.disabled = !isValidKrs;
      engButton.disabled = !isValidKrs;
    });

    plButton.addEventListener('click', () => setLanguageAndRun(0));
    engButton.addEventListener('click', () => setLanguageAndRun(1));
  }
});

async function setLanguageAndRun(languageCode) {
  const krsInputValue = document.getElementById("krsInput").value;
  if (krsInputValue.length !== 10) {
    alert('Proszę wpisać dokładnie 10 cyfr.');
    return;
  }
  console.log(`KRS Number: ${krsInputValue}, Language Code: ${languageCode}`); // Debugging output
  run(krsInputValue, languageCode);
}

async function run(nrKRS, language) {
  try {
    return Word.run(async (context) => {
      
      // Pobranie nr KRS od użytkownika 
      const nrKRS = document.getElementById("krsInput").value

      // Połączenie z API KRS i pobranie odpisu aktualnego i odpisu pełnego
      const krsDataCurrentExcerpt = await fetch(`https://api-krs.ms.gov.pl/api/krs/OdpisAktualny/${nrKRS}?rejestr=P&format=json`);
        
      if (!krsDataCurrentExcerpt.ok) {
        throw new Error(`Error fetching current excerpt: ${krsDataCurrentExcerpt.statusText}`);
      }

      const DataCurrentExcerpt = await krsDataCurrentExcerpt.json();
      
      if (!DataCurrentExcerpt.odpis) {
        throw new Error('Invalid response format for current excerpt.');
      }

      const krsDataFullExcerpt = await fetch(`https://api-krs.ms.gov.pl/api/krs/OdpisPelny/${nrKRS}?rejestr=P&format=json`);
      
      if (!krsDataFullExcerpt.ok) {
        throw new Error(`Error fetching full excerpt: ${krsDataFullExcerpt.statusText}`);
      }

      const DataFullExcerpt = await krsDataFullExcerpt.json();
      
      if (!DataFullExcerpt.odpis) {
        throw new Error('Invalid response format for full excerpt.');
      }

      // Przetworzenie danych z obiektu JSON 
      const companyName = DataCurrentExcerpt.odpis.dane.dzial1.danePodmiotu.nazwa;
      const companyForm = DataCurrentExcerpt.odpis.dane.dzial1.danePodmiotu.formaPrawna;
      let companySeat = DataCurrentExcerpt.odpis.dane.dzial1.siedzibaIAdres.siedziba.miejscowosc;
      let street = DataCurrentExcerpt.odpis.dane.dzial1.siedzibaIAdres.adres.ulica;
      const houseNumber = DataCurrentExcerpt.odpis.dane.dzial1.siedzibaIAdres.adres.nrDomu;
      const postalCode = DataCurrentExcerpt.odpis.dane.dzial1.siedzibaIAdres.adres.kodPocztowy;
      const nrNIP = DataCurrentExcerpt.odpis.dane.dzial1.danePodmiotu.identyfikatory.nip;
      let nrREGON = DataCurrentExcerpt.odpis.dane.dzial1.danePodmiotu.identyfikatory.regon;
      let city = DataCurrentExcerpt.odpis.dane.dzial1.siedzibaIAdres.adres.miejscowosc;
      
      function trimREGON(text) {
        if (text.length > 9) {
            return text.substring(0, 9);
        }
        return text;
    }
      nrREGON = trimREGON(nrREGON);
    
      function formatStreetName(text) {
        const ulIndex = text.search(/ul\.|ulica/i);
        let before, after;
    
        if (ulIndex !== -1) {
            const splitPoint = ulIndex + (text.substring(ulIndex, ulIndex + 4).toUpperCase() === "ULIC" ? 5 : 3);
            before = text.substring(0, splitPoint);
            after = text.substring(splitPoint);
        } else {
            before = '';
            after = text;
        }
    
        // Normalizacja pierwszej części tekstu
        before = before.replace("UL.", "ul.").replace(/Ulica/gi, "ulica");
    
        // Funkcja do zmiany wielkości liter z wyjątkiem określonych słów
        function capitalizeText(text) {
            return text
                .split(' ') // Rozdzielenie tekstu na słowa po spacji
                .map(word => {
                    const lowerCaseExceptions = ["dla", "w", "przy", "pod", "nad", "i"];
                    if (lowerCaseExceptions.includes(word.toLowerCase())) {
                        return word; // Zwróć słowo bez zmian, jeśli jest w wyjątkach
                    }
                    // Kapitalizacja pierwszej litery, reszta liter na małe
                    // Dodatkowo obsługa słów złożonych rozdzielonych myślnikiem
                    return word.split('-').map(part =>
                        part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()
                    ).join('-');
                })
                .join(' '); // Połączenie słów z powrotem w cały tekst
        }
    
        // Stosowanie funkcji do drugiej części tekstu
        after = capitalizeText(after);
    
        // Połączenie części tekstu
        return before + after;
    }

      street = formatStreetName(street);
      
      function capitalizeText(text) {
        const words = text.split(' ');
    
        // Przetwarza każde słowo oddzielnie
        const capitalizedWords = words.map(word => {
            // Rozdziela słowo na człony, jeśli zawiera myślnik
            const parts = word.split('-');
            
            // Kapitalizuje każdy człon słowa
            const capitalizedParts = parts.map(part => {
                if (part.length > 0) {
                    return part[0].toUpperCase() + part.slice(1).toLowerCase();
                }
                return part;
            });
    
            // Łączy przetworzone człony z powrotem, używając myślnika jako separatora
            return capitalizedParts.join('-');
        });
    
        // Łączy przetworzone słowa z powrotem w pełny tekst, używając spacji jako separatora
        return capitalizedWords.join(' ');
      }
      
      city = capitalizeText(city);
      companySeat = capitalizeText(companySeat);

      function capitalizeText(text) {
      // Rozdziela tekst na słowa na podstawie spacji
      const words = text.split(' ');

      // Przetwarza każde słowo oddzielnie
      const capitalizedWords = words.map(word => {
          // Rozdziela słowo na człony, jeśli zawiera myślnik
          const parts = word.split('-');
          
          // Kapitalizuje każdy człon słowa
          const capitalizedParts = parts.map(part => {
              if (part.length > 0) {
                  return part[0].toUpperCase() + part.slice(1).toLowerCase();
              }
              return part;
          });

          // Łączy przetworzone człony z powrotem, używając myślnika jako separatora
          return capitalizedParts.join('-');
      });

      // Łączy przetworzone słowa z powrotem w pełny tekst, używając spacji jako separatora
      return capitalizedWords.join(' ');
      }

      function translateCityName(name) {
        const cityMap = {
            'warszawa': 'Warsaw',
            'kraków': 'Cracow'
        };
    
        return cityMap[name.toLowerCase()] || name;
      }
    
      if (language === 1) {
        companySeat = translateCityName(companySeat);
        city = translateCityName(city);
      }

      //Pozyskanie najnowszego wpisu zawierającego nazwę sądu rejestrowego 
      const recordsCourt = DataFullExcerpt.odpis.naglowekP.wpis;
      
      let court = null;

      for (let i = recordsCourt.length - 1; i >=0; i--) {
        const entry = recordsCourt[i];
        if (entry.oznaczenieSaduDokonujacegoWpisu && entry.oznaczenieSaduDokonujacegoWpisu.includes("SĄD")) {
          court = entry.oznaczenieSaduDokonujacegoWpisu;
          break
        }
      }

      function removeWhitespaceUntilComma(text) {
        // Znajdź indeks pierwszego wystąpienia przecinka
        const commaIndex = text.indexOf(',');
    
        // Jeśli nie ma przecinka, usuń wszystkie białe znaki
        if (commaIndex === -1) {
            return text.replace(/\s+/g, '');
        }
    
        // Podziel tekst na część przed przecinkiem i po przecinku
        const beforeComma = text.slice(0, commaIndex);
        const afterComma = text.slice(commaIndex);
    
        // Usuń białe znaki tylko z części przed przecinkiem
        const cleanedBeforeComma = beforeComma.replace(/\s+/g, '');
    
        // Połącz oczyszczoną część przed przecinkiem z niezmienioną częścią po przecinku
        return cleanedBeforeComma + afterComma;
      }
      
      court = removeWhitespaceUntilComma(court);

      function translateCourtName(courtText, language) {
        let replacements;
    
        if (language === 1) {  // Angielski
            replacements = {
                "SĄDREJONOWYDLAM.ST.WARSZAWYWWARSZAWIE": "District Court for the Capital City of Warsaw in Warsaw",
                "SĄDREJONOWYDLAKRAKOWAŚRÓDMIEŚCIAWKRAKOWIE": "District Court for Kraków Śródmieście in Kraków",
                "SĄDREJONOWYDLAKRAKOWA-ŚRÓDMIEŚCIAWKRAKOWIE": "District Court for Kraków Śródmieście in Kraków",
                "SĄDREJONOWYWBIAŁYMSTOKU": "District Court in Białystok",
                "SĄDREJONOWYWBIELSKU-BIAŁEJ": "in Bielsko-Biała",
                "SĄDREJONOWYWBYDGOSZCZY": "District Court in Bydgoszcz",
                "SĄDREJONOWYWCZĘSTOCHOWIE": "District Court in Częstochowa",
                "SĄDREJONOWYGDAŃSK-PÓŁNOCWGDAŃSKU": "District Court Gdańsk Północ in Gdańsk",
                "SĄDREJONOWYGDAŃSKPÓŁNOCWGDAŃSKU": "District Court Gdańsk Północ in Gdańsk",
                "SĄDREJONOWYDLAWROCŁAWIA-FABRYCZNEJWEWROCŁAWIU": "District Court Wrocław-Fabryczna in Wrocław",
                "SĄDREJONOWYDLAWROCŁAWIAFABRYCZNEJWEWROCŁAWIU": "District Court Wrocław-Fabryczna in Wrocław",
                "SĄDREJONOWYWGLIWICACH": "District Court in Gliwice",
                "SĄDREJONOWYKATOWICE-WSCHÓDWKATOWICACH": "District Court Katowice-Wschód in Katowice",
                "SĄDREJONOWYKATOWICEWSCHÓDWKATOWICACH": "District Court Katowice-Wschód in Katowice",
                "SĄDREJONOWYWKIELCACH": "District Court in Kielce", 
                "SĄDREJONOWYWKOSZALINIE": "District Court in Koszalin",
                "SĄDREJONOWYLUBLINWSCHÓDWLUBLINIEZSIEDZIBĄWŚWIDNIKU": "District Court Lublin-Wschód in Lublin with its registered office in Świdnik",
                "SĄDREJONOWYLUBLIN-WSCHÓDWLUBLINIEZSIEDZIBĄWŚWIDNIKU": "District Court Lublin-Wschód in Lublin with its registered office in Świdnik",
                "SĄDREJONOWYLUBLIN-WSCHÓD": "District Court Lublin-Wschód in Lublin",
                "SĄDREJONOWYLUBLINWSCHÓD": "District Court Lublin-Wschód in Lublin",
                "SĄDREJONOWYDLAŁODZIŚRÓDMIEŚCIAWŁODZI": "District Court for Łódź Śródmieście in Łódź",
                "SĄDREJONOWYDLAŁODZI-ŚRÓDMIEŚCIAWŁODZI": "District Court for Łódź Śródmieście in Łódź",
                "SĄDREJONOWYWOLSZTYNIE": "District Court in Olsztyn",
                "SĄDREJONOWYWOPOLU": "District Court in Opole",
                "SĄDREJONOWYPOZNAŃ-NOWEMIASTOIWILDAWPOZNANIU": "District Court Poznań-Nowe Miasto and Wilda in Poznań",
                "SĄDREJONOWYPOZNAŃNOWEMIASTOIWILDAWPOZNANIU": "District Court Poznań-Nowe Miasto and Wilda in Poznań",
                "SĄDREJONOWYWRZESZOWIE": "District Court in Rzeszów",
                "SĄDREJONOWYSZCZECIN-CENTRUMWSZCZECINIE": "District Court Szczecin-Centrum in Szczecin",
                "SĄDREJONOWYSZCZECINCENTRUMWSZCZECINIE": "District Court Szczecin-Centrum in Szczecin",
                "SĄDREJONOWYWTORUNIU": "District Court in Toruń",
                "SĄDREJONOWYWZIELONEJGÓRZE": "District Court in Zielona Góra",
                "WYDZIAŁ GOSPODARCZY KRAJOWEGO REJESTRU SĄDOWEGO": "Commercial Division of the National Court Register"
            };
        } else if (language === 0) {  // POLSKI
            replacements = {
              "SĄDREJONOWYDLAM.ST.WARSZAWYWWARSZAWIE": "Sąd Rejonowy dla m. st. Warszawy w Warszawie",
              "SĄDREJONOWYDLAKRAKOWAŚRÓDMIEŚCIAWKRAKOWIE": "Sąd Rejonowy dla Krakowa Śródmieścia w Krakowie",
              "SĄDREJONOWYDLAKRAKOWA-ŚRÓDMIEŚCIAWKRAKOWIE": "Sąd Rejonowy dla Krakowa Śródmieścia w Krakowie",
              "SĄDREJONOWYWBIAŁYMSTOKU": "Sąd Rejonowy w Białymstoku",
              "SĄDREJONOWYWBIELSKU-BIAŁEJ": "Sąd Rejonowy w Bielsko-Biała",
              "SĄDREJONOWYWBYDGOSZCZY": "Sąd Rejonowy w Bydgoszczy",
              "SĄDREJONOWYWCZĘSTOCHOWIE": "Sąd Rejonowy w Częstochowie",
              "SĄDREJONOWYGDAŃSK-PÓŁNOCWGDAŃSKU": "Sąd Rejonowy Gdańsk Północ w Gdańsku",
              "SĄDREJONOWYGDAŃSKPÓŁNOCWGDAŃSKU": "Sąd Rejonowy Gdańsk Północ w Gdańsku",
              "SĄDREJONOWYDLAWROCŁAWIA-FABRYCZNEJWEWROCŁAWIU": "Sąd Rejonowy Wrocław-Fabryczna we Wrocławiu",
              "SĄDREJONOWYDLAWROCŁAWIAFABRYCZNEJWEWROCŁAWIU": "Sąd Rejonowy Wrocław-Fabryczna we Wrocławiu",
              "SĄDREJONOWYWGLIWICACH": "Sąd Rejonowy w Gliwicach",
              "SĄDREJONOWYKATOWICE-WSCHÓDWKATOWICACH": "Sąd Rejonowy Katowice-Wschód w Katowicach",
              "SĄDREJONOWYKATOWICEWSCHÓDWKATOWICACH": "Sąd Rejonowy Katowice-Wschód w Katowicach",
              "SĄDREJONOWYWKIELCACH": "Sąd Rejonowy w Kielcach", 
              "SĄDREJONOWYWKOSZALINIE": "Sąd Rejonowy w Koszalinie",
              "SĄDREJONOWYLUBLINWSCHÓDWLUBLINIEZSIEDZIBĄWŚWIDNIKU": "Sąd Rejonowy Lublin-Wschód w Lublinie z siedzibą w Świdniku",
              "SĄDREJONOWYLUBLIN-WSCHÓDWLUBLINIEZSIEDZIBĄWŚWIDNIKU": "Sąd Rejonowy Lublin-Wschód w Lublinie z siedzibą w Świdniku",
              "SĄDREJONOWYLUBLIN-WSCHÓD": "Sąd Rejonowy Lublin-Wschód w Lublinie",
              "SĄDREJONOWYLUBLINWSCHÓD": "Sąd Rejonowy Lublin-Wschód w Lublinie",
              "SĄDREJONOWYDLAŁODZIŚRÓDMIEŚCIAWŁODZI": "Sąd Rejonowy dla Łódzi Śródmieście w Łódzi",
              "SĄDREJONOWYDLAŁODZI-ŚRÓDMIEŚCIAWŁODZI": "Sąd Rejonowy dla Łódzi Śródmieście w Łódzi",
              "SĄDREJONOWYWOLSZTYNIE": "Sąd Rejonowy w Olsztynie",
              "SĄDREJONOWYWOPOLU": "Sąd Rejonowy w Opolu",
              "SĄDREJONOWYPOZNAŃ-NOWEMIASTOIWILDAWPOZNANIU": "Sąd Rejonowy Poznań-Nowe Miasto i Wilda w Poznaniu",
              "SĄDREJONOWYPOZNAŃNOWEMIASTOIWILDAWPOZNANIU": "Sąd Rejonowy Poznań-Nowe Miasto i Wilda w Poznaniu",
              "SĄDREJONOWYWRZESZOWIE": "Sąd Rejonowy w Rzeszowie",
              "SĄDREJONOWYSZCZECIN-CENTRUMWSZCZECINIE": "Sąd Rejonowy Szczecin-Centrum w Szczecinie",
              "SĄDREJONOWYSZCZECINCENTRUMWSZCZECINIE": "Sąd Rejonowy Szczecin-Centrum w Szczecinie",
              "SĄDREJONOWYWTORUNIU": "Sąd Rejonowy w Toruniu",
              "SĄDREJONOWYWZIELONEJGÓRZE": "Sąd Rejonowy w Zielonej Górze",
              "WYDZIAŁ GOSPODARCZY KRAJOWEGO REJESTRU SĄDOWEGO": "Wydział Gospodarczy Krajowego Rejestru Sądowego"
            };
        }
    
        // Iteracja przez obiekt replacements i zamiana odpowiednich fragmentów tekstu
        for (const [key, value] of Object.entries(replacements)) {
            courtText = courtText.replace(new RegExp(key, 'g'), value);
        }
    
        return courtText;
    }

      //OPISZ KEJS PODMIOTÓW WYKREŚLONYCH!!!
    
      court = translateCourtName(court, language);

      //Pozyskanie danych o kapitale zakładowym - dotyczy wyłącznie spółek akcyjnych, z o.o. oraz komandytowo-akcyjnych
      let shareCapital = ""
      let currencyShareCapital = ""
      if (companyForm === 'SPÓŁKA AKCYJNA' || companyForm === 'SPÓŁKA Z OGRANICZONĄ ODPOWIEDZIALNOŚCIĄ' || companyForm === 'SPÓŁKA KOMANDYTOWO-AKCYJNA') {
        shareCapital = DataCurrentExcerpt.odpis.dane.dzial1.kapital.wysokoscKapitaluZakladowego.wartosc;
        currencyShareCapital = DataCurrentExcerpt.odpis.dane.dzial1.kapital.wysokoscKapitaluZakladowego.waluta;
      }
      
      //Pozyskanie danych o kapitale wpłaconym - dotyczy wyłącznie spółek akcyjnych i komandytowo-akcyjnych
      let paidUpCapital = ""
      let paidUpShareCapitalWithCurrency = ""
      if (companyForm === 'SPÓŁKA AKCYJNA' || companyForm === 'SPÓŁKA KOMANDYTOWO-AKCYJNA') {
        paidUpCapital = DataCurrentExcerpt.odpis.dane.dzial1.kapital.czescKapitaluWplaconegoPokrytego.wartosc;
        const paidUpCapitalCurrency = DataCurrentExcerpt.odpis.dane.dzial1.kapital.czescKapitaluWplaconegoPokrytego.waluta;
        if (shareCapital === paidUpCapital) {
          paidUpShareCapitalWithCurrency = ", kapitał wpłacony w całości";
        } else {
        paidUpShareCapitalWithCurrency = ", kapitał wpłacony" + " " + paidUpCapital + " " + paidUpCapitalCurrency
        }
      }

      //Pozyskanie danych o kapitale akcyjnym - dotyczy wyłącznie prostych spółek akcyjnych
      let shareCapitalSJSCwithCurrency = ""
      if (companyForm === 'PROSTA SPÓŁKA AKCYJNA') {
        const shareCapitalSJSC = DataCurrentExcerpt.odpis.dane.dzial1.kapitalPSA.wysokoscKapitaluAkcyjnego.wartosc;
        const shareCapitalCurrencySJSC = DataCurrentExcerpt.odpis.dane.dzial1.kapitalPSA.wysokoscKapitaluAkcyjnego.waluta;
        shareCapitalSJSCwithCurrency = shareCapitalSJSC + " " + shareCapitalCurrencySJSC;
      }

      // Przypisanie wzoru opisu spółki do odpowiedniej spółki i Wstawienie przetworzonych danych do dokumentu Word - WERSJA PL
      let basePatternPL = `${companyName} z siedzibą w ${companySeat}, adres: ${street} ${houseNumber}, ${postalCode} ${city}, wpisaną do Rejestru Przedsiębiorców Krajowego Rejestru Sądowego, dla której ${court} przechowuje akta rejestrowe pod numerem KRS: ${nrKRS}, posiadająca numer identyfikacji podatkowej NIP: ${nrNIP} oraz numer statystyczny REGON: ${nrREGON}`;

      let additionalDetailPL = "";
      if (companyForm === 'SPÓŁKA AKCYJNA' || companyForm === 'SPÓŁKA Z OGRANICZONĄ ODPOWIEDZIALNOŚCIĄ' || companyForm === 'SPÓŁKA KOMANDYTOWO-AKCYJNA') {
          additionalDetailPL = `, o kapitale zakładowym w wysokości ${shareCapital} ${currencyShareCapital}${paidUpShareCapitalWithCurrency} (DALSZE OZNACZENIE PODMIOTU, NP. SPRZEDAJĄCY)`;
      } else if (companyForm === 'PROSTA SPÓŁKA AKCYJNA') {
          additionalDetailPL = `, o kapitale akcyjnym w wysokości ${shareCapitalSJSCwithCurrency} (DALSZE OZNACZENIE PODMIOTU, NP. SPRZEDAJĄCY)`;
      }

      let patternPL = `${basePatternPL}${additionalDetailPL}, reprezentowaną przez:`;
      
      // Przypisanie wzoru opisu spółki do odpowiedniej spółki i wstawienie przetworzonych danych do dokumentu Word - WERSJA ENG

      let basePatternENG = `${companyName} with its registered office in ${companySeat}, address: ${street} ${houseNumber}, ${postalCode} ${city}, entered in the Register of Entrepreneurs of the National Court Register, for which the ${court} maintains the registration files under the KRS number: ${nrKRS}, having the tax identification number (NIP): ${nrNIP} and the statistical number (REGON): ${nrREGON}`;

      let additionalDetailENG = "";
      if (companyForm === 'SPÓŁKA AKCYJNA' || companyForm === 'SPÓŁKA Z OGRANICZONĄ ODPOWIEDZIALNOŚCIĄ' || companyForm === 'SPÓŁKA KOMANDYTOWO-AKCYJNA') {
          additionalDetailENG = `, with the share capital of ${shareCapital} ${currencyShareCapital}${paidUpShareCapitalWithCurrency} (FURTHER INDICATION OF THE ENTITY, E.G. SELLER)`;
      } else if (companyForm === 'PROSTA SPÓŁKA AKCYJNA') {
          additionalDetailENG = `, with the share capital of ${shareCapitalSJSCwithCurrency} (FURTHER INDICATION OF THE ENTITY, E.G. SELLER)`;
      }

      let patternENG = `${basePatternENG}${additionalDetailENG}, represented by:`;

      const range = context.document.getSelection();

      if (language === 0) {
        range.insertText(patternPL, Word.InsertLocation.replace);
      } else {
        range.insertText(patternENG, Word.InsertLocation.replace);
      }

      await context.sync();
    });
  } catch (error) {
    console.error('An error occurred:', error);
    alert(`Wystąpił błąd: ${error.message}`);
  }
}
//komentarz