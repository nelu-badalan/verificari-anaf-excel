# 📊 Verificare Automată CUI-uri ANAF (Excel VBA)

Acest instrument profesional construit în Excel automatizează procesul de verificare a statusului fiscal pentru liste de companii (CUI-uri). Soluția interoghează direct API-ul oficial al ANAF (v9), oferind o alternativă rapidă și eficientă la verificările manuale.
Se pot crea liste custom de cuie, si poate folosi functia isValidCUI, pentru pastrarea celor corecte, deci putem avea CUIE și fără să știm firma.

## 🚀 Funcționalități Principale

- [cite_start]**Procesare Batch (200 CUI):** Codul este optimizat să trimită cereri în pachete de câte 200 de CUI-uri simultan, respectând limita tehnică a API-ului ANAF pentru o viteză maximă[cite: 1].
- **Analiză Comparativă Automată:** Extrage și pune față în față datele de la **sfârșitul lunii precedente** și **sfârșitul lunii anterioare acesteia**, facilitând auditul rapid al partenerilor.
- [cite_start]**Logare Debug și Erori:** Include un modul de Debug care înregistrează pașii procesării și un sheet dedicat pentru erori, unde sunt evidențiate CUI-urile invalide sau problemele de comunicare cu serverul[cite: 1, 19].
- [cite_start]**Parsare JSON Integrată:** Folosește un parser JSON robust (VBA-JSON) pentru a interpreta răspunsurile complexe primite de la server[cite: 24].

## 🛠️ Instrucțiuni de Utilizare

### 1. Pregătirea Fișierului
- Descarcă fișierul `.xlsm` din repository.
- **Important:** Click dreapta pe fișier -> **Properties** -> Bifează **Unblock** -> OK. Această setare Windows este necesară pentru a permite rularea codului VBA descărcat de pe internet.

### 2. Introducerea Datelor
- Deschide fișierul și navighează la sheet-ul principal.
- [cite_start]Introdu lista de CUI-uri (doar numerele, fără prefixul "RO") în coloana **A**, începând cu celula **A2**[cite: 35].
- Asigură-te că nu există rânduri libere între CUI-uri.

### 3. Execuția
- Apasă butonul **„Aduce date ANAF”**.
- [cite_start]Programul va interoga API-ul (URL: `https://webservicesp.anaf.ro/api/PlatitorTvaRest/v9/tva`)[cite: 1].
- Rezultatele vor fi populate automat, iar comparația va fi generată într-un sheet separat.

## 📋 Detalii Tehnice (Pentru Programatori)

- **Limbaj:** VBA (Visual Basic for Applications).
- [cite_start]**Endpoint API:** Versiunea 9 a API-ului de plătitori TVA[cite: 1].
- [cite_start]**Module incluse:** - `Interogare_ANAF`: Modulul principal de control[cite: 24].
    - [cite_start]`JsonConverter`: Responsabil pentru transformarea textului JSON în dicționare și colecții VBA[cite: 24].
- [cite_start]**Debug Mode:** Funcția `EnsureDebug` creează automat un sheet "Debug" unde poți urmări în timp real răspunsurile API-ului în caz de probleme[cite: 1].

## ⚖️ Licență și Limitări

Acest proiect este licențiat sub **MIT License**. 
- **Disclaimer:** Autorul nu este responsabil pentru acuratețea datelor furnizate de API-ul ANAF sau pentru deciziile fiscale luate pe baza acestui fișier. Verificați întotdeauna rezultatele critice pe portalul oficial.

## 👤 Autor

**Nelu Bădălan** [![LinkedIn](https://img.shields.io/badge/LinkedIn-Profile-blue?style=flat-square&logo=linkedin)](https://www.linkedin.com/in/nelu-badalan-8ab7a120/)
