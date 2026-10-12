# Pubblicazione gratuita su Streamlit Community Cloud — v53

Questa guida è pensata per il telefono. **Non avviare la pubblicazione finché il pacchetto non è stato collaudato e la struttura del repository non è stata controllata.** Non è necessario modificare GitHub adesso.

## Perché questa piattaforma

Streamlit Community Cloud è progettato per eseguire applicazioni Streamlit e collega un repository GitHub a un URL `streamlit.app`. Il piano gratuito può mettere a riposo un'app dopo 12 ore senza traffico; il primo accesso successivo può richiedere di risvegliarla. Risorse e condizioni possono cambiare.

Documentazione ufficiale:
- Avvio della pubblicazione: https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy
- Dipendenze Python e Linux (`requirements.txt` e `packages.txt`): https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies
- Limiti e sospensione per inattività: https://docs.streamlit.io/deploy/streamlit-community-cloud/manage-your-app

## Sicurezza di avvio

La versione v54 parte in modalità produzione per impostazione predefinita: se `APP_ACCESS_PASSWORD` non è configurata tra i secrets, l’app si arresta invece di mostrare un’interfaccia accessibile senza password. Configura il secret prima di avviare il servizio.

## Prima della pubblicazione

1. Conservare una copia di sicurezza del repository attuale.
2. Caricare i file del pacchetto nella cartella radice del repository scelto: `app.py`, `requirements.txt`, `packages.txt`, `.streamlit/config.toml` e la cartella `templates/` con tutti gli 8 DOCX. Non caricare carte d'identità, dati di persone o file di prova personali.
3. Verificare che il repository contenga i file di questa versione e non una vecchia cartella annidata.
4. In `https://share.streamlit.io/`, accedere con GitHub e scegliere **Create app**.
5. Selezionare il repository, il ramo e il percorso `app.py` (nella radice del repository, se i file sono stati caricati lì).
6. Aprire **Advanced settings** e nel campo **Secrets** inserire il contenuto seguente, sostituendo il testo di esempio con una password lunga e unica:

```toml
APP_ACCESS_PASSWORD = "SOSTITUISCI_CON_UNA_PASSWORD_LUNGA_E_UNICA"
```

7. Scegliere una versione Python supportata, preferibilmente Python 3.12 se compatibile con il runtime proposto, e avviare il deploy.
8. Attendere la build. Se fallisce, non fare tentativi casuali: aprire i log, copiare il messaggio d'errore senza segreti e correggere la causa.
9. Solo dopo un avvio riuscito, collaudare con immagini e dati completamente fittizi: password, caricamento immagini, correzione dei campi OCR, selezione degli otto modelli, download DOCX e PDF (se LibreOffice è stato installato correttamente), resa su smartphone.

## Privacy e sicurezza: condizione essenziale

Streamlit Community Cloud ospita attualmente le app negli Stati Uniti. Le carte d'identità contengono dati personali: **non usare documenti reali o pratiche amministrative reali su questa piattaforma gratuita prima di una valutazione adeguata della protezione dei dati, del trattamento extra-UE, delle condizioni contrattuali, dei log e della cancellazione dei file**. La password condivisa è solo un controllo iniziale e non identifica singolarmente gli operatori. Il repository può essere pubblico: il codice e i modelli diventano visibili, ma non devono contenere dati personali o segreti.

## Limiti noti

- La build remota e l'interfaccia non sono state ancora provate.
- Il pacchetto apt per LibreOffice può aumentare tempi e risorse di build; il PDF va considerato opzionale finché non è verificato nel cloud.
- L'OCR dipende dalla qualità delle foto; ogni campo estratto deve essere controllato e corretto da una persona.
- Non considerare l'app certificata per l'uso amministrativo fino al collaudo end-to-end.
- Non modificare il repository GitHub prima di aver confermato la struttura finale dei file e aver conservato il codice attuale.
