/**
 * Deutsche Tool-Inhalte für SEO
 * Enthält detaillierte Beschreibungen, Anleitungen, Anwendungsfälle und FAQs für alle Tools
 */

import { ToolContent } from '@/types/tool';

/**
 * Deutsche Tool-Inhalts-Map
 * Jedes Tool enthält: title, metaDescription, keywords, description, howToUse, useCases, faq
 */
export const toolContentDe: Record<string, ToolContent> = {
  // ==================== POPULÄRE TOOLS ====================
  'pdf-multi-tool': {
    title: 'PDF Multi-Tool',
    metaDescription: 'All-in-One PDF-Editor: PDF zusammenfügen, teilen, organisieren, Seiten löschen, drehen und extrahieren in einem Tool.',
    keywords: ['pdf multi tool', 'pdf editor deutsch', 'pdf zusammenfügen', 'pdf teilen', 'pdf organisieren', 'alles in einem pdf'],
    description: `
      <p>Das PDF Multi-Tool ist Ihre Komplettlösung für alle Aufgaben der PDF-Seitenverwaltung. Dieses leistungsstarke All-in-One-Tool kombiniert mehrere PDF-Operationen in einer intuitiven Benutzeroberfläche und spart Ihnen Zeit und Mühe.</p>
      <p>Ob Sie mehrere Dokumente zusammenfügen, ein großes PDF in kleinere Dateien aufteilen, Seiten neu organisieren, unerwünschte Inhalte löschen, Seiten drehen oder bestimmte Abschnitte extrahieren müssen – dieses Tool erledigt alles, ohne dass Sie zwischen verschiedenen Anwendungen wechseln müssen.</p>
      <p>Die gesamte Verarbeitung findet direkt in Ihrem Browser statt, was die Privatsphäre und Sicherheit Ihrer Dokumente gewährleistet. Es werden keine Dateien auf einen Server hochgeladen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Ziehen Sie Ihre PDF-Datei per Drag & Drop in den Upload-Bereich oder klicken Sie, um Dateien von Ihrem Gerät auszuwählen.' },
      { step: 2, title: 'Operation wählen', description: 'Wählen Sie aus den verfügbaren Operationen: Zusammenfügen, Teilen, Organisieren, Seiten löschen, Drehen oder Extrahieren.' },
      { step: 3, title: 'Optionen konfigurieren', description: 'Passen Sie die spezifischen Einstellungen an, wie z. B. Seitenbereiche, Drehwinkel oder die Reihenfolge beim Zusammenfügen.' },
      { step: 4, title: 'Verarbeiten und Herunterladen', description: 'Klicken Sie auf die Schaltfläche zum Verarbeiten und laden Sie Ihre modifizierte PDF-Datei herunter.' },
    ],
    useCases: [
      { title: 'Dokumentenvorbereitung', description: 'Bereiten Sie Unterlagen für die Einreichung vor, indem Sie unnötige Seiten entfernen und mehrere Dateien kombinieren.', icon: 'file-check' },
      { title: 'Berichtserstellung', description: 'Fügen Sie verschiedene Berichtsabschnitte zusammen und organisieren Sie Kapitel zu einem professionellen Gesamtdokument.', icon: 'book-open' },
      { title: 'Archivverwaltung', description: 'Teilen Sie große Archivdateien in handliche Abschnitte auf und extrahieren Sie relevante Seiten.', icon: 'archive' },
    ],
    faq: [
      { question: 'Wie viele PDFs kann ich gleichzeitig verarbeiten?', answer: 'Sie können bis zu 10 PDF-Dateien gleichzeitig hochladen und verarbeiten, mit einer maximalen Gesamtgröße von 500 MB.' },
      { question: 'Bleiben meine Lesezeichen erhalten?', answer: 'Ja, beim Zusammenfügen von PDFs bleiben vorhandene Lesezeichen erhalten und können optional in einer einheitlichen Struktur kombiniert werden.' },
      { question: 'Gibt es ein Seitenlimit?', answer: 'Es gibt kein striktes Seitenlimit. Das Tool kann Dokumente mit hunderten von Seiten verarbeiten, wobei sehr große Dateien etwas länger dauern können.' },
    ],
  },

  'merge-pdf': {
    title: 'PDF zusammenfügen',
    metaDescription: 'Kombinieren Sie mehrere PDF-Dateien zu einem Dokument. Kostenloser Online-PDF-Merger mit Drag-and-Drop-Sortierung.',
    keywords: ['pdf zusammenfügen', 'pdf kombinieren', 'pdf verbinden', 'pdf merger deutsch'],
    description: `
      <p>Mit "PDF zusammenfügen" können Sie mehrere PDF-Dokumente schnell und einfach zu einer einzigen Datei kombinieren. Ob Sie Berichte konsolidieren, gescannte Dokumente verbinden oder eine Präsentation zusammenstellen – dieses Tool macht den Prozess nahtlos.</p>
      <p>Laden Sie einfach Ihre Dateien hoch, ordnen Sie sie per Drag & Drop in der gewünschten Reihenfolge an und fügen Sie sie zusammen. Das Tool bewahrt die Qualität Ihrer Originaldateien.</p>
    `,
    howToUse: [
      { step: 1, title: 'Dateien hochladen', description: 'Ziehen Sie mehrere PDF-Dateien in den Bereich oder wählen Sie sie manuell aus.' },
      { step: 2, title: 'Reihenfolge anordnen', description: 'Verschieben Sie die Miniaturansichten, um die gewünschte Abfolge festzulegen.' },
      { step: 3, title: 'Zusammenfügen', description: 'Klicken Sie auf "Zusammenfügen" und laden Sie das fertige Dokument herunter.' },
    ],
    useCases: [
      { title: 'Berichte kombinieren', description: 'Fügen Sie Monats- oder Quartalsberichte zu einem Jahresdokument zusammen.', icon: 'file-text' },
      { title: 'Portfolios erstellen', description: 'Kombinieren Sie Arbeitsproben, Zertifikate und Anschreiben zu einem professionellen Portfolio.', icon: 'briefcase' },
      { title: 'Rechnungen bündeln', description: 'Fassen Sie mehrere Belege oder Rechnungen für die Buchhaltung in einer Datei zusammen.', icon: 'receipt' },
    ],
    faq: [
      { question: 'Wie viele PDFs kann ich verbinden?', answer: 'Sie können bis zu 100 PDF-Dateien auf einmal mit einer Gesamtgröße von bis zu 500 MB zusammenfügen.' },
      { question: 'Bleibt die Qualität erhalten?', answer: 'Ja, der Prozess erfolgt ohne Qualitätsverlust oder zusätzliche Kompression.' },
      { question: 'Kann ich passwortgeschützte PDFs zusammenfügen?', answer: 'Diese müssen zuerst entschlüsselt werden. Nutzen Sie dafür unser Tool "PDF entschlüsseln".' },
    ],
  },

  'rotate-custom': {
    title: 'Eigener Drehwinkel',
    metaDescription: 'PDF-Seiten um jeden beliebigen Winkel drehen. Präzise Ausrichtung für schief gescannte Dokumente.',
    keywords: ['pdf drehen winkel', 'pdf begradigen', 'pdf schief gescannt', 'pdf rotation'],
    description: `
      <p>Dieses Tool gibt Ihnen präzise Kontrolle über die Ausrichtung Ihrer PDF-Seiten. Im Gegensatz zu Standard-Tools, die nur 90-Grad-Schritte unterstützen, können Sie hier jeden spezifischen Winkel eingeben.</p>
      <p>Ideal zum Begradigen von Dokumenten, die schräg eingezogen wurden, oder zum Anpassen von technischen Zeichnungen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Laden Sie die PDF-Datei hoch, deren Seiten gedreht werden müssen.' },
      { step: 2, title: 'Winkel einstellen', description: 'Geben Sie den exakten Grad für die Drehung ein.' },
      { step: 3, title: 'Vorschau und Download', description: 'Prüfen Sie die Ausrichtung in der Echtzeit-Vorschau und speichern Sie das Ergebnis.' },
    ],
    useCases: [
      { title: 'Gescannte Dokumente', description: 'Begradigen Sie Seiten, die schräg durch den Scanner gelaufen sind.', icon: 'scan' },
      { title: 'Technische Zeichnungen', description: 'Passen Sie die Ausrichtung von Plänen präzise an.', icon: 'ruler' },
      { title: 'Kreative Layouts', description: 'Erstellen Sie künstlerische Layouts durch individuelle Drehung.', icon: 'pen-tool' },
    ],
    faq: [
      { question: 'Sind Dezimalzahlen möglich?', answer: 'Aktuell unterstützen wir ganzzahlige Gradangaben, arbeiten aber an einer Unterstützung für Dezimalstellen.' },
      { question: 'Wird der Inhalt abgeschnitten?', answer: 'Nein, die Seitengröße wird automatisch angepasst, damit der gedrehte Inhalt vollständig sichtbar bleibt.' },
      { question: 'Kann ich nur eine einzelne Seite drehen?', answer: 'Ja, Sie können für jede Seite einen individuellen Winkel festlegen.' },
    ],
  },

  'split-pdf': {
    title: 'PDF teilen',
    metaDescription: 'PDF-Dateien in mehrere Dokumente aufteilen. Extrahieren Sie einzelne Seiten oder teilen Sie nach Bereichen.',
    keywords: ['pdf teilen', 'pdf trennen', 'pdf seiten extrahieren', 'pdf splitter deutsch'],
    description: `
      <p>Mit "PDF teilen" können Sie ein einzelnes PDF in mehrere kleinere Dateien aufteilen. Perfekt, um Kapitel zu extrahieren oder kombinierte Dokumente wieder zu trennen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das PDF aus, das Sie aufteilen möchten.' },
      { step: 2, title: 'Methode wählen', description: 'Wählen Sie zwischen Seitenbereichen, Einzel-Extraktion oder Aufteilung in festen Intervallen.' },
      { step: 3, title: 'Teilen und Speichern', description: 'Klicken Sie auf "Teilen" und laden Sie die Ergebnisse (ggf. als ZIP) herunter.' },
    ],
    useCases: [
      { title: 'Kapitel extrahieren', description: 'Teilen Sie ein Buch in einzelne Kapitel auf.', icon: 'book' },
      { title: 'Sammelscans trennen', description: 'Trennen Sie einen Stapelscan in die ursprünglichen Einzeldokumente.', icon: 'copy' },
      { title: 'Handouts erstellen', description: 'Extrahieren Sie nur die relevanten Folien einer Präsentation.', icon: 'presentation' },
    ],
    faq: [
      { question: 'Kann ich das PDF in jede einzelne Seite zerlegen?', answer: 'Ja, wählen Sie die Option "Jede Seite einzeln speichern".' },
      { question: 'Bleiben Lesezeichen erhalten?', answer: 'Lesezeichen, die auf die extrahierten Seiten verweisen, bleiben in der neuen Datei bestehen.' },
      { question: 'Wie erhalte ich die Dateien?', answer: 'Bei mehreren Dateien werden diese bequem in einem ZIP-Archiv zusammengefasst.' },
    ],
  },

  'compress-pdf': {
    title: 'PDF komprimieren',
    metaDescription: 'PDF-Dateigröße reduzieren bei gleichbleibender Qualität. Online-PDF-Kompressor für kleinere Dateien.',
    keywords: ['pdf komprimieren', 'pdf verkleinern', 'pdf größe reduzieren', 'pdf optimieren'],
    description: `
      <p>Dieses Tool reduziert die Dateigröße Ihrer PDFs, ideal für E-Mail-Anhänge oder Web-Uploads. Sie können zwischen verschiedenen Kompressionsstufen wählen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie die Datei aus, die verkleinert werden soll.' },
      { step: 2, title: 'Stufe wählen', description: 'Wählen Sie: Niedrig (Beste Qualität), Mittel (Ausbalanciert) oder Hoch (Kleinste Datei).' },
      { step: 3, title: 'Komprimieren', description: 'Starten Sie den Vorgang und laden Sie die optimierte PDF herunter.' },
    ],
    useCases: [
      { title: 'E-Mail-Anhänge', description: 'Unterschreiten Sie Größenlimits von Mail-Anbietern.', icon: 'mail' },
      { title: 'Web-Veröffentlichung', description: 'Schnellere Ladezeiten für Dokumente auf Ihrer Website.', icon: 'globe' },
      { title: 'Speicherplatz sparen', description: 'Archivieren Sie Dokumente platzsparend auf Ihrer Festplatte.', icon: 'hard-drive' },
    ],
    faq: [
      { question: 'Wie stark wird die Datei verkleinert?', answer: 'Das hängt vom Inhalt ab. Bilder können oft um 50-80% verkleinert werden, reiner Text weniger.' },
      { question: 'Leidet die Textqualität?', answer: 'Nein, Text bleibt scharf. Die Kompression wirkt sich primär auf Bilder und Grafiken aus.' },
      { question: 'Ist die Nutzung sicher?', answer: 'Ja, die Kompression erfolgt lokal in Ihrem Browser; Ihre Daten verlassen Ihr Gerät nicht.' },
    ],
  },

  'edit-pdf': {
    title: 'PDF bearbeiten',
    metaDescription: 'PDF-Dateien online bearbeiten. Text, Bilder, Anmerkungen und Formen hinzufügen.',
    keywords: ['pdf bearbeiten', 'pdf editor online', 'pdf beschriften', 'text in pdf einfügen'],
    description: `
      <p>Unser PDF-Editor bietet Ihnen Werkzeuge zum Ändern und Kommentieren Ihrer Dokumente – ganz ohne teure Software. Fügen Sie Text, Bilder, Formen und Markierungen hinzu.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das zu bearbeitende Dokument aus.' },
      { step: 2, title: 'Werkzeug wählen', description: 'Nutzen Sie die Toolbar für Text, Highlights, Formen oder Bilder.' },
      { step: 3, title: 'Änderungen vornehmen', description: 'Platzieren und gestalten Sie die Elemente direkt auf dem PDF.' },
      { step: 4, title: 'Speichern', description: 'Laden Sie die bearbeitete Version herunter.' },
    ],
    useCases: [
      { title: 'Dokumentenprüfung', description: 'Kommentieren und markieren Sie Entwürfe im Team.', icon: 'message-square' },
      { title: 'Formulare ausfüllen', description: 'Beschriften Sie PDFs und fügen Sie Unterschriften hinzu.', icon: 'edit-3' },
      { title: 'Schwärzen', description: 'Überdecken Sie sensible Informationen vor der Weitergabe.', icon: 'eye-off' },
    ],
    faq: [
      { question: 'Kann ich den Originaltext löschen?', answer: 'Dieses Tool dient primär dem Hinzufügen von Inhalten. Zum direkten Löschen von Originaltext ist oft das Quelldokument (z.B. Word) nötig.' },
      { question: 'Sind die Änderungen dauerhaft?', answer: 'Ja, nach dem Speichern werden die Anmerkungen fest in die PDF-Ebenen integriert.' },
      { question: 'Gibt es eine Rückgängig-Funktion?', answer: 'Ja, während der Bearbeitung können Sie Schritte jederzeit rückgängig machen.' },
    ],
  },

  'jpg-to-pdf': {
    title: 'JPG in PDF',
    metaDescription: 'Konvertieren Sie JPG-Bilder in PDF. Mehrere JPG-Dateien zu einem PDF-Dokument zusammenfassen.',
    keywords: ['jpg in pdf', 'jpeg in pdf', 'bilder in pdf umwandeln', 'foto zu pdf'],
    description: `
      <p>Wandeln Sie Ihre JPEG-Bilder schnell in PDF-Dokumente um. Sie können einzelne Fotos oder ganze Bildserien konvertieren und die Seitenreihenfolge anpassen.</p>
    `,
    howToUse: [
      { step: 1, title: 'Bilder hochladen', description: 'Wählen Sie eine oder mehrere JPG-Dateien aus.' },
      { step: 2, title: 'Anordnen', description: 'Bringen Sie die Bilder per Drag & Drop in die richtige Reihenfolge.' },
      { step: 3, title: 'Konvertieren', description: 'Erstellen Sie das PDF und laden Sie es herunter.' },
    ],
    useCases: [
      { title: 'Fotoalben', description: 'Erstellen Sie ein PDF-Album aus Urlaubs- oder Eventfotos.', icon: 'image' },
      { title: 'Belege digitalisieren', description: 'Wandeln Sie Handyfotos von Quittungen in saubere PDFs um.', icon: 'camera' },
      { title: 'Portfolio', description: 'Fassen Sie Design-Arbeiten in einem kompakten Dokument zusammen.', icon: 'folder' },
    ],
    faq: [
      { question: 'Wie viele Bilder sind möglich?', answer: 'Sie können bis zu 100 Bilder in ein einzelnes PDF umwandeln.' },
      { question: 'Bleibt die Bildqualität hoch?', answer: 'Ja, die Bilder werden in ihrer Originalqualität eingebettet.' },
      { question: 'Kann ich die Seitengröße wählen?', answer: 'Ja, Sie können zwischen Originalgröße, A4 oder US-Letter wählen.' },
    ],
  },

  'sign-pdf': {
    title: 'PDF unterschreiben',
    metaDescription: 'Elektronische Unterschriften zu PDF-Dokumenten hinzufügen. Zeichnen, tippen oder Signatur hochladen.',
    keywords: ['pdf unterschreiben', 'elektronische signatur', 'e-signatur', 'pdf signieren online'],
    description: `
      <p>Fügen Sie Ihre Unterschrift sicher und schnell zu PDFs hinzu. Zeichnen Sie Ihre Signatur mit der Maus/Touchpad, tippen Sie sie ein oder laden Sie ein Bild Ihrer Unterschrift hoch.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das zu unterzeichnende Dokument.' },
      { step: 2, title: 'Signatur erstellen', description: 'Zeichnen, tippen oder laden Sie Ihre Unterschrift hoch.' },
      { step: 3, title: 'Platzieren', description: 'Klicken Sie an die Stelle im PDF, an der die Signatur erscheinen soll.' },
    ],
    useCases: [
      { title: 'Verträge signieren', description: 'Unterschreiben Sie Verträge ohne lästiges Drucken und Scannen.', icon: 'file-signature' },
      { title: 'Formulare', description: 'Bestätigen Sie Anträge oder Einverständniserklärungen digital.', icon: 'clipboard' },
      { title: 'Freigabeprozesse', description: 'Geben Sie Dokumente im beruflichen Umfeld schnell frei.', icon: 'check-circle' },
    ],
    faq: [
      { question: 'Ist die Signatur rechtsgültig?', answer: 'In den meisten Ländern sind elektronische Signaturen für Standardverträge rechtlich bindend. Für notarielle Dokumente gelten Sonderregeln.' },
      { question: 'Wird meine Signatur gespeichert?', answer: 'Sie können sie optional lokal im Browser speichern, um sie beim nächsten Mal direkt wiederzuverwenden.' },
      { question: 'Kann ich mehrere Stellen signieren?', answer: 'Ja, Sie können die Signatur beliebig oft auf verschiedenen Seiten platzieren.' },
    ],
  },

  'crop-pdf': {
    title: 'PDF zuschneiden',
    metaDescription: 'PDF-Seiten zuschneiden, um Ränder zu entfernen. PDF-Dokumente präzise trimmen.',
    keywords: ['pdf zuschneiden', 'pdf ränder entfernen', 'pdf trimmen', 'pdf format anpassen'],
    description: `
      <p>Entfernen Sie weiße Ränder oder unerwünschte Bereiche von Ihren PDF-Seiten. Ideal, um den Fokus auf den wesentlichen Inhalt zu legen oder Seitenformate zu vereinheitlichen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das PDF aus, das Sie zuschneiden möchten.' },
      { step: 2, title: 'Bereich wählen', description: 'Ziehen Sie den Rahmen auf die gewünschte Größe.' },
      { step: 3, title: 'Anwenden', description: 'Wählen Sie aus, ob alle Seiten oder nur bestimmte Seiten beschnitten werden sollen.' },
    ],
    useCases: [
      { title: 'Ränder entfernen', description: 'Trimmen Sie überflüssigen Platz bei Scans.', icon: 'maximize-2' },
      { title: 'Inhalt fokussieren', description: 'Entfernen Sie Kopf- oder Fußzeilen für eine bessere Lesbarkeit.', icon: 'target' },
      { title: 'Format-Korrektur', description: 'Bringen Sie alle Seiten auf eine einheitliche Größe.', icon: 'square' },
    ],
    faq: [
      { question: 'Geht der Inhalt verloren?', answer: 'Ja, alles außerhalb des gewählten Rahmens wird entfernt. Behalten Sie das Original als Backup.' },
      { question: 'Kann ich jede Seite anders zuschneiden?', answer: 'Ja, Sie können für jede Seite oder Seitengruppen individuelle Schnittmasken festlegen.' },
      { question: 'Bleibt die Textqualität gleich?', answer: 'Ja, da nur der Sichtbereich geändert wird, bleibt die Qualität der Vektoren und Bilder unberührt.' },
    ],
  },

  'extract-pages': {
    title: 'Seiten extrahieren',
    metaDescription: 'Bestimmte Seiten aus einer PDF extrahieren und als neues Dokument speichern.',
    keywords: ['pdf seiten extrahieren', 'seiten aus pdf speichern', 'pdf teilauszug'],
    description: `
      <p>Wählen Sie gezielt Seiten aus einem großen Dokument aus und erstellen Sie daraus eine neue, kompakte Datei. Ideal für Auszüge aus Büchern oder Berichten.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das Quelldokument aus.' },
      { step: 2, title: 'Seiten wählen', description: 'Klicken Sie auf die Miniaturansichten der Seiten, die Sie behalten möchten.' },
      { step: 3, title: 'Extrahieren', description: 'Klicken Sie auf "Extrahieren" und laden Sie das neue PDF herunter.' },
    ],
    useCases: [
      { title: 'Auszüge erstellen', description: 'Speichern Sie nur relevante Kapitel eines Handbuchs.', icon: 'file-minus' },
      { title: 'Gezielte Weitergabe', description: 'Teilen Sie nur die Seiten, die für den Empfänger wichtig sind.', icon: 'share-2' },
      { title: 'Archivierung', description: 'Sichern Sie nur die wichtigsten Seiten eines langen Dokuments.', icon: 'archive' },
    ],
    faq: [
      { question: 'Kann ich nicht-aufeinanderfolgende Seiten wählen?', answer: 'Ja, Sie können beliebige Seiten im Dokument anklicken, egal an welcher Stelle sie stehen.' },
      { question: 'Bleiben Links in der Datei aktiv?', answer: 'Ja, interne und externe Links auf den extrahierten Seiten bleiben funktionsfähig.' },
      { question: 'Wird das Original verändert?', answer: 'Nein, es wird eine neue Datei erstellt; Ihr Original bleibt unberührt.' },
    ],
  },

  'organize-pdf': {
    title: 'PDF organisieren',
    metaDescription: 'Seitenreihenfolge ändern, Seiten duplizieren oder löschen. PDF-Dokumente einfach neu strukturieren.',
    keywords: ['pdf organisieren', 'pdf seiten sortieren', 'pdf reihenfolge ändern'],
    description: `
      <p>Strukturieren Sie Ihre PDF-Dokumente völlig neu. Mit einer einfachen Drag-and-Drop-Oberfläche können Sie Seiten verschieben, löschen oder wichtige Abschnitte duplizieren.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie die Datei aus, die Sie neu ordnen möchten.' },
      { step: 2, title: 'Sortieren', description: 'Verschieben Sie die Seiten per Maus. Nutzen Sie die Buttons zum Löschen oder Kopieren.' },
      { step: 3, title: 'Speichern', description: 'Laden Sie das organisierte Dokument herunter.' },
    ],
    useCases: [
      { title: 'Scan-Fehler korrigieren', description: 'Bringen Sie falsch herum eingescannte Seiten in die richtige Reihenfolge.', icon: 'arrow-up-down' },
      { title: 'Eigene Struktur', description: 'Erstellen Sie eine individuelle Abfolge für Präsentationen.', icon: 'list' },
      { title: 'Bereinigen', description: 'Entfernen Sie Leerseiten oder doppelte Inhalte sofort.', icon: 'trash-2' },
    ],
    faq: [
      { question: 'Kann ich Seiten duplizieren?', answer: 'Ja, jede Seite kann mit einem Klick kopiert und an eine andere Stelle verschoben werden.' },
      { question: 'Gibt es eine Vorschau?', answer: 'Ja, Sie sehen große Vorschaubilder aller Seiten, was das Sortieren erleichtert.' },
      { question: 'Bleibt die Dateigröße gleich?', answer: 'Meistens ja, außer Sie fügen viele Duplikate hinzu oder löschen viele Seiten.' },
    ],
  },

  'delete-pages': {
    title: 'Seiten löschen',
    metaDescription: 'Unerwünschte Seiten aus PDF-Dateien entfernen. Einfaches Auswählen und Löschen.',
    keywords: ['pdf seiten löschen', 'pdf seiten entfernen', 'seiten aus pdf ausschneiden'],
    description: `
      <p>Entfernen Sie schnell und unkompliziert nicht benötigte Seiten aus Ihren PDFs. Ob Leerseiten, veraltete Inhalte oder sensible Daten – mit diesem Tool bereinigen Sie Ihre Dokumente in Sekunden.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das Dokument aus, das bereinigt werden soll.' },
      { step: 2, title: 'Seiten wählen', description: 'Markieren Sie die Seiten, die entfernt werden sollen.' },
      { step: 3, title: 'Löschen und Speichern', description: 'Bestätigen Sie das Löschen und laden Sie die gekürzte PDF herunter.' },
    ],
    useCases: [
      { title: 'Leerseiten entfernen', description: 'Bereinigen Sie Dokumente von versehentlich mitgescannten Leerseiten.', icon: 'file-x' },
      { title: 'Datenschutz', description: 'Löschen Sie Seiten mit vertraulichen Infos, bevor Sie den Rest teilen.', icon: 'shield' },
      { title: 'Veraltete Inhalte', description: 'Entfernen Sie nicht mehr aktuelle Abschnitte aus Handbüchern.', icon: 'filter' },
    ],
    faq: [
      { question: 'Ist das Löschen endgültig?', answer: 'In der heruntergeladenen Datei ja. Ihr lokales Original auf dem PC bleibt jedoch unverändert.' },
      { question: 'Kann ich mehrere Seiten gleichzeitig löschen?', answer: 'Ja, Sie können beliebig viele Seiten markieren und in einem Schritt entfernen.' },
      { question: 'Was passiert mit dem Inhaltsverzeichnis?', answer: 'Die Seitenzahlen im Dokument bleiben meist gleich, aber Links zu gelöschten Seiten funktionieren nicht mehr.' },
    ],
  },

  'ocr-pdf': {
    title: 'PDF mit OCR erkennen',
    metaDescription: 'Gescannte PDFs durchsuchbar machen. Text aus Bildern und Scans extrahieren.',
    keywords: ['pdf ocr', 'texterkennung pdf', 'pdf durchsuchbar machen', 'scan in text'],
    description: `
      <p>Verwandeln Sie statische Scans in intelligente Dokumente. Unsere OCR-Technologie erkennt Text in Bildern und fügt eine unsichtbare, durchsuchbare Textebene hinzu.</p>
    `,
    howToUse: [
      { step: 1, title: 'Scan hochladen', description: 'Wählen Sie ein bildbasiertes PDF aus.' },
      { step: 2, title: 'Sprache wählen', description: 'Wählen Sie die Sprache des Dokuments für optimale Ergebnisse.' },
      { step: 3, title: 'OCR starten', description: 'Laden Sie das nun durchsuchbare und kopierbare PDF herunter.' },
    ],
    useCases: [
      { title: 'Archive digitalisieren', description: 'Machen Sie alte Papierarchive per Volltextsuche findbar.', icon: 'archive' },
      { title: 'Text kopieren', description: 'Kopieren Sie Text aus Dokumenten, die nur als Bild vorliegen.', icon: 'type' },
      { title: 'Barrierefreiheit', description: 'Machen Sie Scans für Screenreader lesbar.', icon: 'accessibility' },
    ],
    faq: [
      { question: 'Welche Sprachen werden unterstützt?', answer: 'Über 100 Sprachen, darunter Deutsch, Englisch, Französisch und Spanisch.' },
      { question: 'Wie gut ist die Erkennung?', answer: 'Bei sauberen Scans liegt die Genauigkeit oft über 98%.' },
      { question: 'Bleibt das Aussehen gleich?', answer: 'Ja, das visuelle Erscheinungsbild bleibt exakt identisch.' },
    ],
  },

  'pdf-to-docx': {
    title: 'PDF in Word',
    metaDescription: 'PDF in editierbare Word-Dokumente (DOCX) umwandeln. Layout und Formatierung bleiben erhalten.',
    keywords: ['pdf in word', 'pdf zu docx', 'pdf umwandeln word', 'pdf editierbar machen'],
    description: `
      <p>Konvertieren Sie PDF-Dokumente in vollständig editierbare Microsoft Word-Dateien. Layouts, Schriftarten und Bilder werden so originalgetreu wie möglich übernommen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF wählen', description: 'Laden Sie das Dokument hoch.' },
      { step: 2, title: 'Konvertierung', description: 'Warten Sie kurz, während die Struktur analysiert wird.' },
      { step: 3, title: 'Word-Datei laden', description: 'Bearbeiten Sie den Text nun direkt in Word oder Google Docs.' },
    ],
    useCases: [
      { title: 'Verträge anpassen', description: 'PDF-Verträge in Word ändern und neu verhandeln.', icon: 'file-text' },
      { title: 'Lebenslauf aktualisieren', description: 'Bringen Sie alte PDF-Lebensläufe ohne Tipparbeit auf den neuesten Stand.', icon: 'user' },
      { title: 'Datenübernahme', description: 'Übernehmen Sie komplexe Texte in Ihre eigene Textverarbeitung.', icon: 'copy' },
    ],
    faq: [
      { question: 'Wird die Formatierung übernommen?', answer: 'Wir geben unser Bestes, Spalten, Tabellen und Bilder exakt zu platzieren.' },
      { question: 'Funktioniert das bei Scans?', answer: 'Für Scans empfehlen wir zuerst unser OCR-Tool für bessere Textergebnisse.' },
      { question: 'Ist das Ergebnis mit Google Docs kompatibel?', answer: 'Ja, die erstellte DOCX-Datei lässt sich problemlos dort öffnen.' },
    ],
  },
  // ==================== EDIT & ANNOTATE (Fortsetzung) ====================
  'bookmark': {
    title: 'Lesezeichen bearbeiten',
    metaDescription: 'PDF-Lesezeichen hinzufügen, bearbeiten und verwalten. Erstellen Sie eine Navigationsstruktur für Ihre Dokumente.',
    keywords: ['pdf lesezeichen', 'lesezeichen bearbeiten', 'pdf navigation', 'inhaltsverzeichnis erstellen'],
    description: `
      <p>Mit "Lesezeichen bearbeiten" können Sie Lesezeichen in Ihren PDF-Dokumenten erstellen, ändern und organisieren. Lesezeichen ermöglichen eine schnelle Navigation zu bestimmten Abschnitten, was die Nutzung langer Dokumente erheblich erleichtert.</p>
      <p>Sie können neue Lesezeichen hinzufügen, die Hierarchie neu organisieren oder Lesezeichen aus externen Quellen importieren. Dieses Tool ist unerlässlich für die Erstellung professioneller, navigierbarer Dokumente.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie die PDF-Datei aus, deren Lesezeichen Sie verwalten möchten.' },
      { step: 2, title: 'Lesezeichen verwalten', description: 'Fügen Sie neue Punkte hinzu, benennen Sie bestehende um oder ziehen Sie diese per Drag & Drop in eine hierarchische Ordnung.' },
      { step: 3, title: 'Speichern und Download', description: 'Übernehmen Sie die Änderungen und laden Sie das PDF mit der neuen Navigationsstruktur herunter.' },
    ],
    useCases: [
      { title: 'Navigation erstellen', description: 'Helfen Sie Lesern, sich in langen Dokumenten schnell zurechtzufinden.', icon: 'navigation' },
      { title: 'Kapitel organisieren', description: 'Spiegeln Sie die Kapitelstruktur Ihres Dokuments in den Lesezeichen wider.', icon: 'book-open' },
      { title: 'Barrierefreiheit verbessern', description: 'Machen Sie Dokumente benutzerfreundlicher und zugänglicher.', icon: 'accessibility' },
    ],
    faq: [
      { question: 'Kann ich verschachtelte Lesezeichen erstellen?', answer: 'Ja, Sie können eine Baumstruktur mit Haupt- und Unterlesezeichen erstellen.' },
      { question: 'Werden die Lesezeichen überall angezeigt?', answer: 'Ja, Lesezeichen sind ein PDF-Standard und werden von allen gängigen Readern und Browsern unterstützt.' },
      { question: 'Kann ich Lesezeichen importieren?', answer: 'Ja, das Tool unterstützt den Import von Strukturen aus JSON- oder Textdateien.' },
    ],
  },

  'table-of-contents': {
    title: 'Inhaltsverzeichnis erstellen',
    metaDescription: 'Generieren Sie ein Inhaltsverzeichnis für Ihr PDF. Erstellen Sie anklickbare Navigation aus Lesezeichen.',
    keywords: ['pdf inhaltsverzeichnis', 'toc generator', 'pdf index', 'dokumentennavigation'],
    description: `
      <p>Dieses Tool generiert eine navigierbare Inhaltsverzeichnisseite für Ihre PDF-Dokumente. Das Verzeichnis kann aus vorhandenen Lesezeichen oder benutzerdefinierten Einträgen erstellt werden.</p>
      <p>Passen Sie das Erscheinungsbild mit verschiedenen Stilen und Layouts an. Das generierte Verzeichnis enthält anklickbare Links, die direkt zu den entsprechenden Seiten springen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das Dokument aus, für das ein Inhaltsverzeichnis benötigt wird.' },
      { step: 2, title: 'TOC konfigurieren', description: 'Wählen Sie Stil, Schriftart und Position. Entscheiden Sie, ob Lesezeichen als Basis dienen sollen.' },
      { step: 3, title: 'Generieren', description: 'Erstellen Sie das Inhaltsverzeichnis und laden Sie die aktualisierte PDF herunter.' },
    ],
    useCases: [
      { title: 'Akademische Arbeiten', description: 'Fügen Sie Abschlussarbeiten oder Forschungsberichten ein professionelles Verzeichnis hinzu.', icon: 'graduation-cap' },
      { title: 'Geschäftsberichte', description: 'Erstellen Sie übersichtliche Berichte für Stakeholder mit klarer Sektionsauflistung.', icon: 'bar-chart' },
      { title: 'Benutzerhandbücher', description: 'Generieren Sie umfassende Indizes für technische Dokumentationen.', icon: 'book' },
    ],
    faq: [
      { question: 'Ist das Verzeichnis interaktiv?', answer: 'Ja, jeder Eintrag ist ein Link, der den Leser direkt zur Zielseite führt.' },
      { question: 'Wo wird das Verzeichnis eingefügt?', answer: 'Standardmäßig am Anfang des Dokuments, Sie können den Ort aber individuell festlegen.' },
      { question: 'Kann ich das Design anpassen?', answer: 'Ja, es stehen verschiedene Layout-Vorlagen zur Verfügung.' },
    ],
  },

  'page-numbers': {
    title: 'Seitenzahlen hinzufügen',
    metaDescription: 'Seitenzahlen zu PDF-Dokumenten hinzufügen. Position, Format und Startnummer individuell anpassen.',
    keywords: ['pdf seitenzahlen', 'pdf paginierung', 'seiten nummerieren', 'pdf bearbeiten'],
    description: `
      <p>Fügen Sie Ihren PDFs professionelle Seitenzahlen hinzu. Wählen Sie aus verschiedenen Formaten, Positionen und Stilen, um das Layout Ihres Dokuments zu perfektionieren.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Laden Sie das Dokument hoch, das nummeriert werden soll.' },
      { step: 2, title: 'Formatierung wählen', description: 'Bestimmen Sie Position (z.B. unten rechts), Startnummer und das Zahlenformat.' },
      { step: 3, title: 'Anwenden', description: 'Fügen Sie die Paginierung hinzu und laden Sie das Ergebnis herunter.' },
    ],
    useCases: [
      { title: 'Professionelle Berichte', description: 'Stellen Sie sicher, dass Ihre Business-Dokumente eine korrekte Paginierung haben.', icon: 'file-text' },
      { title: 'Rechtliche Dokumente', description: 'Fügen Sie Verträgen für eine bessere Referenzierung Seitenzahlen hinzu.', icon: 'scale' },
      { title: 'Skripte und Manuskripte', description: 'Organisieren Sie Ihre Entwürfe durch eine durchgehende Nummerierung.', icon: 'graduation-cap' },
    ],
    faq: [
      { question: 'Kann ich die erste Seite auslassen?', answer: 'Ja, Sie können festlegen, dass die Nummerierung erst ab einer bestimmten Seite (z.B. nach dem Deckblatt) beginnt.' },
      { question: 'Welche Formate gibt es?', answer: 'Arabische Zahlen (1, 2, 3), römische Zahlen (i, ii, iii) oder das Format "Seite X von Y".' },
      { question: 'Kann ich die Schriftart ändern?', answer: 'Ja, Sie können Schriftgröße und Typ an Ihr Dokument anpassen.' },
    ],
  },

  'add-watermark': {
    title: 'Wasserzeichen hinzufügen',
    metaDescription: 'Text- oder Bild-Wasserzeichen zu PDFs hinzufügen. Schützen und branden Sie Ihre Dokumente.',
    keywords: ['pdf wasserzeichen', 'pdf schützen', 'pdf stempeln', 'branding pdf'],
    description: `
      <p>Schützen Sie Ihr geistiges Eigentum, indem Sie Text- oder Bild-Wasserzeichen auf Ihre PDFs setzen. Ideal für Statusanzeigen wie "Entwurf" oder zur Kennzeichnung mit Ihrem Firmenlogo.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das zu schützende Dokument aus.' },
      { step: 2, title: 'Wasserzeichen erstellen', description: 'Geben Sie Text ein oder laden Sie ein Logo hoch. Passen Sie Transparenz und Winkel an.' },
      { step: 3, title: 'Platzieren und Speichern', description: 'Wenden Sie das Wasserzeichen auf alle oder ausgewählte Seiten an.' },
    ],
    useCases: [
      { title: 'Urheberschutz', description: 'Versehen Sie Bilder oder Dokumente mit einem Copyright-Vermerk.', icon: 'copyright' },
      { title: 'Status-Kennzeichnung', description: 'Markieren Sie Dokumente deutlich als "KOPIE", "ENTWURF" oder "VERTRAULICH".', icon: 'shield' },
      { title: 'Corporate Identity', description: 'Integrieren Sie Ihr Firmenlogo dezent im Hintergrund Ihrer Unterlagen.', icon: 'award' },
    ],
    faq: [
      { question: 'Kann ich die Transparenz einstellen?', answer: 'Ja, Sie können das Wasserzeichen fast unsichtbar oder deckend einstellen.' },
      { question: 'Werden alle Seiten markiert?', answer: 'Sie können wählen: Alle Seiten, nur die erste Seite oder ein benutzerdefinierter Bereich.' },
      { question: 'Kann ich Bilder (PNG/JPG) nutzen?', answer: 'Ja, sowohl Text als auch Bilddateien (inkl. Transparenz) werden unterstützt.' },
    ],
  },

  'header-footer': {
    title: 'Kopf- & Fußzeile',
    metaDescription: 'Kopf- und Fußzeilen zu PDF hinzufügen. Seitenzahlen, Daten und eigenen Text einfügen.',
    keywords: ['pdf kopfzeile', 'pdf fußzeile', 'pdf beschriften', 'briefpapier pdf'],
    description: `
      <p>Erstellen Sie ein einheitliches Layout für Ihre PDFs, indem Sie Kopf- und Fußzeilen hinzufügen. Fügen Sie dynamische Felder wie das aktuelle Datum, den Dateinamen oder Seitenzahlen ein.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie Ihr Dokument aus.' },
      { step: 2, title: 'Inhalt definieren', description: 'Geben Sie Text für die linke, mittlere oder rechte Position in Kopf/Fußzeile ein.' },
      { step: 3, title: 'Layout anpassen', description: 'Wählen Sie Abstände zum Rand und Schriftstile aus.' },
    ],
    useCases: [
      { title: 'Geschäftskorrespondenz', description: 'Fügen Sie Firmennamen und Kontaktdaten zu jedem PDF hinzu.', icon: 'briefcase' },
      { title: 'Juristische Schriftsätze', description: 'Fügen Sie Aktenzeichen und Seitenzahlen standardisiert ein.', icon: 'scale' },
      { title: 'Dokumentation', description: 'Kennzeichnen Sie Versionen und Daten in der Kopfzeile.', icon: 'file-text' },
    ],
    faq: [
      { question: 'Sind unterschiedliche Kopfzeilen möglich?', answer: 'Ja, Sie können für gerade und ungerade Seiten verschiedene Inhalte festlegen.' },
      { question: 'Kann ich Variablen nutzen?', answer: 'Ja, Felder wie [Datum] oder [Dateiname] werden automatisch ausgefüllt.' },
      { question: 'Überdeckt das den Inhalt?', answer: 'Sie können die Ränder des Dokuments anpassen, um Platz für die Zeilen zu schaffen.' },
    ],
  },

  // ==================== VISUALS & COLORS ====================
  'invert-colors': {
    title: 'Farben invertieren',
    metaDescription: 'PDF-Farben für den Dark Mode invertieren. Dokumente in Negativ-Farben umwandeln.',
    keywords: ['pdf farben invertieren', 'pdf dark mode', 'pdf negativ', 'augen schonen pdf'],
    description: `
      <p>Invertieren Sie die Farben Ihrer PDF-Dokumente, um einen Negativ-Effekt zu erzielen. Dies ist besonders nützlich für das Lesen in dunkler Umgebung (Dark Mode), um die Augen zu schonen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Laden Sie das PDF hoch, das Sie farblich umkehren möchten.' },
      { step: 2, title: 'Optionen wählen', description: 'Entscheiden Sie, ob auch Bilder invertiert werden sollen oder nur Text/Hintergrund.' },
      { step: 3, title: 'Invertieren', description: 'Laden Sie die augenschonende Version Ihres PDFs herunter.' },
    ],
    useCases: [
      { title: 'Nachtmodus', description: 'Angenehmeres Lesen von hellen Dokumenten bei Nacht.', icon: 'moon' },
      { title: 'Barrierefreiheit', description: 'Hilfe für Nutzer mit Sehschwäche durch höheren Kontrast.', icon: 'eye' },
      { title: 'Tinte sparen', description: 'Invertieren Sie dunkle Dokumente mit viel schwarzem Hintergrund vor dem Drucken.', icon: 'printer' },
    ],
    faq: [
      { question: 'Sieht das Dokument dann aus wie ein Negativ-Foto?', answer: 'Ja, weißer Hintergrund wird schwarz und schwarzer Text wird weiß.' },
      { question: 'Bleiben Bilder normal?', answer: 'Sie können wählen, ob Bilder von der Invertierung ausgeschlossen werden sollen.' },
      { question: 'Ist das umkehrbar?', answer: 'Ja, Sie können das resultierende PDF einfach erneut invertieren, um die Originalfarben fast exakt wiederherzustellen.' },
    ],
  },// ==================== VISUALS & COLORS (Fortsetzung) ====================
  'background-color': {
    title: 'Hintergrundfarbe ändern',
    metaDescription: 'PDF-Hintergrundfarbe ändern. Fügen Sie farbige Hintergründe zu Ihren Dokumentseiten hinzu.',
    keywords: ['pdf hintergrundfarbe', 'pdf farbe ändern', 'farbiges pdf', 'seitenfarbe anpassen'],
    description: `
      <p>Mit diesem Tool können Sie die Hintergrundfarbe Ihrer PDF-Seiten ändern oder eine neue hinzufügen. Dies kann die Lesbarkeit verbessern, das Dokument optisch aufwerten oder es an Ihr Branding anpassen.</p>
      <p>Wählen Sie eine beliebige Farbe aus und wenden Sie diese auf das gesamte Dokument oder nur auf bestimmte Seiten an. Alle vorhandenen Inhalte bleiben dabei im Vordergrund erhalten.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das Dokument aus, dessen Hintergrund Sie einfärben möchten.' },
      { step: 2, title: 'Farbe wählen', description: 'Nutzen Sie den Farbwähler oder geben Sie einen Hex-Code ein.' },
      { step: 3, title: 'Anwenden', description: 'Speichern Sie das PDF mit dem neuen Hintergrund-Layer.' },
    ],
    useCases: [
      { title: 'Lesbarkeit verbessern', description: 'Nutzen Sie ein sanftes Beige oder Sepia, um die Augen bei langem Lesen zu entlasten.', icon: 'eye' },
      { title: 'Markendesign', description: 'Passen Sie Präsentationsunterlagen an Ihre Firmenfarben an.', icon: 'palette' },
      { title: 'Bereiche markieren', description: 'Verwenden Sie unterschiedliche Farben, um Kapitel optisch zu trennen.', icon: 'layers' },
    ],
    faq: [
      { question: 'Wird der Text überdeckt?', answer: 'Nein, die Farbe wird als unterste Ebene hinzugefügt, sodass Text und Bilder sichtbar bleiben.' },
      { question: 'Kann ich verschiedene Farben pro Seite nutzen?', answer: 'Ja, Sie können das Tool für einzelne Seiten oder Bereiche separat anwenden.' },
      { question: 'Kann ich vorhandene Farben entfernen?', answer: 'Dieses Tool fügt Farben hinzu. Zum Entfernen komplexer Hintergründe nutzen Sie bitte den PDF-Editor.' },
    ],
  },

  'text-color': {
    title: 'Textfarbe ändern',
    metaDescription: 'Ändern Sie die Textfarbe in PDF-Dokumenten. Modifizieren Sie die Farbe aller Textinhalte zentral.',
    keywords: ['pdf textfarbe ändern', 'text umfärben pdf', 'kontrast verbessern pdf'],
    description: `
      <p>Passen Sie die Farbe aller Textelemente in Ihrem PDF an. Dies ist ideal, um den Kontrast zu erhöhen, ein Dokument für den Druck vorzubereiten oder die Optik an Corporate-Design-Vorgaben anzupassen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Laden Sie das PDF-Dokument hoch.' },
      { step: 2, title: 'Farbe festlegen', description: 'Wählen Sie die neue Farbe für alle enthaltenen Texte aus.' },
      { step: 3, title: 'Konvertieren', description: 'Das Tool färbt alle Schriften um und stellt das PDF zum Download bereit.' },
    ],
    useCases: [
      { title: 'Kontrast optimieren', description: 'Ändern Sie hellen Text in Schwarz, um die Lesbarkeit zu garantieren.', icon: 'contrast' },
      { title: 'Einheitliches Branding', description: 'Färben Sie Texte in Ihre spezifische Markenfarbe um.', icon: 'palette' },
      { title: 'Barrierefreiheit', description: 'Passen Sie Farben an, um Web-Standards für Kontraste zu erfüllen.', icon: 'accessibility' },
    ],
    faq: [
      { question: 'Wird wirklich jeder Text geändert?', answer: 'Ja, das Tool erkennt Textelemente und wendet die neue Farbe global an.' },
      { question: 'Bleiben Fettdruck und Kursivschrift erhalten?', answer: 'Ja, alle Formatierungen bleiben bestehen, nur der Farbwert wird geändert.' },
      { question: 'Gilt das auch für Bilder?', answer: 'Nein, Text in Bildern (Rastergrafiken) kann auf diese Weise nicht umgefärbt werden.' },
    ],
  },

  'add-stamps': {
    title: 'Stempel hinzufügen',
    metaDescription: 'Stempel zu PDF-Dokumenten hinzufügen. Nutzen Sie Vorlagen für Genehmigungen, Entwürfe und mehr.',
    keywords: ['pdf stempel', 'genehmigt stempel pdf', 'pdf abstempeln', 'digitaler stempel'],
    description: `
      <p>Versehen Sie Ihre PDFs mit digitalen Stempeln. Nutzen Sie klassische Bürostempel wie "GENEHMIGT", "ABGELEHNT" oder "ENTWURF", oder laden Sie Ihr eigenes Stempelbild hoch.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das zu stempelnde Dokument.' },
      { step: 2, title: 'Stempel wählen', description: 'Wählen Sie eine Vorlage oder laden Sie ein eigenes PNG/JPG hoch.' },
      { step: 3, title: 'Platzieren', description: 'Klicken Sie auf die gewünschte Stelle, passen Sie die Größe an und speichern Sie.' },
    ],
    useCases: [
      { title: 'Freigabeprozesse', description: 'Markieren Sie Rechnungen oder Verträge als "Geprüft" oder "Bezahlt".', icon: 'check-circle' },
      { title: 'Status-Updates', description: 'Kennzeichnen Sie Dokumente als "Final" oder "Veraltet".', icon: 'tag' },
      { title: 'Qualitätskontrolle', description: 'Fügen Sie Prüfsiegel oder Inspektionsstempel hinzu.', icon: 'clipboard-check' },
    ],
    faq: [
      { question: 'Welche Vorlagen gibt es?', answer: 'Genehmigt, Abgelehnt, Entwurf, Vertraulich, Kopie und viele mehr.' },
      { question: 'Kann ich eigene Logos nutzen?', answer: 'Ja, Sie können jedes Bild als individuellen Stempel hochladen.' },
      { question: 'Kann ich mehrere Stempel nutzen?', answer: 'Ja, Sie können so viele Stempel wie nötig auf verschiedenen Seiten platzieren.' },
    ],
  },

  'remove-annotations': {
    title: 'Anmerkungen entfernen',
    metaDescription: 'Anmerkungen aus PDF-Dateien löschen. Bereinigen Sie Kommentare, Highlights und Markups.',
    keywords: ['pdf anmerkungen löschen', 'kommentare entfernen pdf', 'pdf bereinigen', 'markups entfernen'],
    description: `
      <p>Entfernen Sie alle Kommentare, Markierungen und Notizen aus Ihrem PDF. So erstellen Sie eine saubere Version des Dokuments für die finale Veröffentlichung oder Weitergabe.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Laden Sie das kommentierte Dokument hoch.' },
      { step: 2, title: 'Typen wählen', description: 'Wählen Sie aus, ob nur Kommentare, Highlights oder alle Markups gelöscht werden sollen.' },
      { step: 3, title: 'Bereinigen', description: 'Laden Sie das "saubere" PDF ohne Korrekturspuren herunter.' },
    ],
    useCases: [
      { title: 'Finalisierung', description: 'Entfernen Sie interne Korrekturhinweise vor dem Versand an Kunden.', icon: 'file-check' },
      { title: 'Datenschutz', description: 'Löschen Sie potenziell sensible Kommentare aus dem Review-Prozess.', icon: 'shield' },
      { title: 'Saubere Kopie', description: 'Erstellen Sie eine Druckversion ohne störende Hervorhebungen.', icon: 'copy' },
    ],
    faq: [
      { question: 'Welche Elemente werden gelöscht?', answer: 'Highlights, Unterstreichungen, Haftnotizen, Stempel und Freihandzeichnungen.' },
      { question: 'Bleibt der Text erhalten?', answer: 'Ja, nur die darüberliegenden Anmerkungen werden entfernt; der eigentliche Inhalt bleibt unberührt.' },
      { question: 'Ist das rückgängig zu machen?', answer: 'Nach dem Download nicht mehr. Behalten Sie für den Fall der Fälle Ihr Original.' },
    ],
  },

  // ==================== FORM TOOLS ====================
  'form-filler': {
    title: 'Formular-Ausfüller',
    metaDescription: 'PDF-Formulare online ausfüllen. Bearbeiten Sie interaktive Formulare direkt im Browser.',
    keywords: ['pdf formular ausfüllen', 'pdf form filler', 'interaktives pdf bearbeiten'],
    description: `
      <p>Füllen Sie interaktive PDF-Formulare direkt im Browser aus. Geben Sie Text ein, setzen Sie Häkchen in Checkboxen und wählen Sie Optionen aus Dropdown-Menüs – ganz ohne Drucken.</p>
    `,
    howToUse: [
      { step: 1, title: 'Formular hochladen', description: 'Wählen Sie die PDF-Formulardatei aus.' },
      { step: 2, title: 'Ausfüllen', description: 'Klicken Sie in die Felder und geben Sie Ihre Daten ein.' },
      { step: 3, title: 'Speichern', description: 'Laden Sie das fertig ausgefüllte Formular herunter.' },
    ],
    useCases: [
      { title: 'Anträge & Behörden', description: 'Füllen Sie Anmeldeformulare oder Anträge digital aus.', icon: 'clipboard' },
      { title: 'Steuerformulare', description: 'Bearbeiten Sie Finanzdokumente bequem am Rechner.', icon: 'file-text' },
      { title: 'Vertragsdaten', description: 'Ergänzen Sie Ihre persönlichen Daten in Vertragsentwürfen.', icon: 'file-signature' },
    ],
    faq: [
      { question: 'Kann ich den Fortschritt speichern?', answer: 'Ja, Sie können das teilweise ausgefüllte Formular speichern und später weiterbearbeiten.' },
      { question: 'Was ist "Flattening"?', answer: 'Dabei werden die Felder in statischen Text umgewandelt, damit sie nach dem Versand nicht mehr geändert werden können.' },
      { question: 'Werden XFA-Formulare unterstützt?', answer: 'Ja, das Tool unterstützt sowohl Standard-AcroForms als auch XFA-Formate.' },
    ],
  },

  'form-creator': {
    title: 'Formular-Ersteller',
    metaDescription: 'Erstellen Sie ausfüllbare PDF-Formulare. Fügen Sie Textfelder, Checkboxen und Dropdowns hinzu.',
    keywords: ['pdf formular erstellen', 'ausfüllbares pdf machen', 'pdf formfelder hinzufügen'],
    description: `
      <p>Verwandeln Sie statische PDFs in interaktive Formulare. Fügen Sie Textfelder, Checkboxen, Radio-Buttons und Dropdowns hinzu, um professionelle Datenerfassung zu ermöglichen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Laden Sie das Dokument hoch, das als Basis für Ihr Formular dient.' },
      { step: 2, title: 'Felder hinzufügen', description: 'Wählen Sie Feldtypen aus der Toolbar und platzieren Sie diese per Klick.' },
      { step: 3, title: 'Konfigurieren', description: 'Legen Sie Eigenschaften wie Pflichtfelder fest und speichern Sie das Ergebnis.' },
    ],
    useCases: [
      { title: 'Bewerbungsbögen', description: 'Erstellen Sie strukturierte Formulare für Job-Interessenten.', icon: 'user-plus' },
      { title: 'Umfragen', description: 'Bauen Sie interaktive Fragebögen zur Datenerhebung auf.', icon: 'clipboard-list' },
      { title: 'Bestellformulare', description: 'Erstellen Sie Dokumente mit Mengenfeldern und Auswahloptionen.', icon: 'shopping-cart' },
    ],
    faq: [
      { question: 'Welche Feldtypen gibt es?', answer: 'Textzeilen, Checkboxen, Auswahlknöpfe, Listen, Datumsfelder und Signaturfelder.' },
      { question: 'Sind Berechnungen möglich?', answer: 'Ja, einfache Summen- oder Durchschnittsberechnungen können für Zahlenfelder konfiguriert werden.' },
      { question: 'Können Felder Pflichtfelder sein?', answer: 'Ja, Sie können festlegen, dass bestimmte Felder ausgefüllt werden müssen.' },
    ],
  },

  'remove-blank-pages': {
    title: 'Leerseiten entfernen',
    metaDescription: 'Erkennen und löschen Sie automatisch leere Seiten aus Ihren PDF-Dokumenten.',
    keywords: ['pdf leerseiten löschen', 'leere seiten entfernen pdf', 'pdf bereinigen'],
    description: `
      <p>Dieses intelligente Tool erkennt und entfernt automatisch leere Seiten aus Ihren Dokumenten. Ideal zum Bereinigen von Scans oder zum Entfernen von Trennblättern nach dem Zusammenfügen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das zu bereinigende Dokument.' },
      { step: 2, title: 'Empfindlichkeit wählen', description: 'Stellen Sie ein, ab wann eine Seite als "leer" gilt (z.B. trotz kleiner Scan-Punkte).' },
      { step: 3, title: 'Entfernen', description: 'Laden Sie das kompakte PDF ohne unnötige Leerseiten herunter.' },
    ],
    useCases: [
      { title: 'Scan-Optimierung', description: 'Entfernen Sie Rückseiten von Dokumenten, die nur einseitig bedruckt waren.', icon: 'scan' },
      { title: 'Dateigröße reduzieren', description: 'Sparen Sie Platz, indem Sie nutzlose Seiten löschen.', icon: 'minimize-2' },
      { title: 'Trennseiten löschen', description: 'Entfernen Sie Platzhalterseiten aus zusammengefügten Stapeln.', icon: 'minus' },
    ],
    faq: [
      { question: 'Wie funktioniert die Erkennung?', answer: 'Das Tool analysiert den Inhalt der Seite. Seiten mit fast keinem sichtbaren Inhalt werden als leer markiert.' },
      { question: 'Was ist mit Schmutzpartikeln auf dem Scan?', answer: 'Sie können den Schwellenwert anpassen, sodass auch Seiten mit minimalen "Rausch-Punkten" als leer erkannt werden.' },
      { question: 'Sehe ich vorher, was gelöscht wird?', answer: 'Ja, die erkannten Seiten werden in einer Vorschau markiert, bevor Sie das Löschen bestätigen.' },
    ],
  },// ==================== CONVERT TO PDF ====================
  'image-to-pdf': {
    title: 'Bild in PDF',
    metaDescription: 'Konvertieren Sie beliebige Bilder in PDF. Unterstützung für JPG, PNG, WebP, BMP, TIFF, SVG und HEIC.',
    keywords: ['bild in pdf', 'foto zu pdf umwandeln', 'bilder kombinieren pdf', 'grafik zu pdf'],
    description: `
      <p>Wandeln Sie Bilder jeglicher Formate in professionelle PDF-Dokumente um. Mit Unterstützung für JPG, PNG, WebP, BMP, TIFF, SVG und HEIC ist dies Ihr universeller Bildkonverter.</p>
      <p>Kombinieren Sie mehrere Bilder zu einer einzigen PDF-Datei, ordnen Sie diese nach Belieben an und passen Sie Seitengröße sowie Ausrichtung individuell an.</p>
    `,
    howToUse: [
      { step: 1, title: 'Bilder hochladen', description: 'Ziehen Sie Bilder in den Upload-Bereich oder wählen Sie Dateien von Ihrem Gerät.' },
      { step: 2, title: 'Anordnen & Konfigurieren', description: 'Sortieren Sie die Bilder und legen Sie das Seitenformat (z. B. A4 oder Originalgröße) fest.' },
      { step: 3, title: 'Konvertieren', description: 'Erstellen Sie Ihr PDF und laden Sie das Ergebnis sofort herunter.' },
    ],
    useCases: [
      { title: 'Fotosammlungen', description: 'Fassen Sie Urlaubsfotos oder Event-Bilder in einem Album-PDF zusammen.', icon: 'images' },
      { title: 'Dokumenten-Archiv', description: 'Digitalisieren Sie Papierunterlagen, indem Sie Bild-Scans in PDF archivieren.', icon: 'archive' },
      { title: 'Portfolios', description: 'Erstellen Sie eine professionelle Mappe aus Ihren Design-Entwürfen.', icon: 'file-stack' },
    ],
    faq: [
      { question: 'Welche Formate werden unterstützt?', answer: 'JPG, PNG, WebP, BMP, TIFF, SVG und das Apple-Format HEIC.' },
      { question: 'Bleibt die Bildqualität erhalten?', answer: 'Ja, Bilder werden standardmäßig in ihrer Originalauflösung eingebettet.' },
      { question: 'Kann ich die Reihenfolge ändern?', answer: 'Ja, Sie können die Bilder per Drag & Drop sortieren, bevor Sie das PDF generieren.' },
    ],
  },

  'png-to-pdf': {
    title: 'PNG in PDF',
    metaDescription: 'Konvertieren Sie PNG-Bilder in PDF. Erhalten Sie Transparenzen und kombinieren Sie mehrere PNGs.',
    keywords: ['png in pdf', 'png umwandeln', 'transparente bilder pdf', 'screenshot zu pdf'],
    description: `
      <p>Konvertieren Sie PNG-Dateien in PDF, während Transparenzen erhalten bleiben. Ideal für Grafiken, Logos und Screenshots, die einen transparenten Hintergrund besitzen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PNGs hochladen', description: 'Wählen Sie Ihre PNG-Dateien aus.' },
      { step: 2, title: 'Layout wählen', description: 'Bestimmen Sie die Seitengröße und die Bildanordnung.' },
      { step: 3, title: 'Herunterladen', description: 'Laden Sie das fertige PDF-Dokument herunter.' },
    ],
    useCases: [
      { title: 'Grafik-Portfolios', description: 'Präsentieren Sie Logos und UI-Designs in einem sauberen Dokument.', icon: 'palette' },
      { title: 'Software-Dokumentation', description: 'Fassen Sie Programm-Screenshots zu einer Anleitung zusammen.', icon: 'monitor' },
      { title: 'Logo-Kataloge', description: 'Erstellen Sie eine Übersicht Ihrer Marken-Assets.', icon: 'award' },
    ],
    faq: [
      { question: 'Bleibt die Transparenz erhalten?', answer: 'Ja, transparente Bereiche im PNG werden im PDF korrekt dargestellt.' },
      { question: 'Was passiert mit animierten PNGs?', answer: 'Diese werden als statisches Bild (das erste Frame) konvertiert.' },
      { question: 'Kann ich die Hintergrundfarbe wählen?', answer: 'Ja, Sie können für transparente Bereiche eine Hintergrundfarbe im PDF festlegen.' },
    ],
  },

  'webp-to-pdf': {
    title: 'WebP in PDF',
    metaDescription: 'WebP-Bilder in PDF konvertieren. Modernes Bildformat einfach für Druck und Archivierung umwandeln.',
    keywords: ['webp in pdf', 'google bildformat konvertieren', 'webp zu pdf deutsch'],
    description: `
      <p>Wandeln Sie moderne WebP-Bilder in das universell kompatible PDF-Format um. Ideal, um Bilder direkt aus dem Web für den Druck oder die Langzeitarchivierung vorzubereiten.</p>
    `,
    howToUse: [
      { step: 1, title: 'WebP hochladen', description: 'Wählen Sie WebP-Dateien aus Ihrem Ordner aus.' },
      { step: 2, title: 'Optionen anpassen', description: 'Wählen Sie Hoch- oder Querformat für Ihre PDF-Seiten.' },
      { step: 3, title: 'Speichern', description: 'Generieren Sie das PDF aus Ihren WebP-Grafiken.' },
    ],
    useCases: [
      { title: 'Web-Inhalte archivieren', description: 'Speichern Sie Bilder von Webseiten dauerhaft im PDF-Format.', icon: 'globe' },
      { title: 'Druckvorbereitung', description: 'Machen Sie moderne Web-Bilder für Standarddrucker verfügbar.', icon: 'printer' },
      { title: 'Format-Standardisierung', description: 'Konvertieren Sie WebP in das plattformübergreifende PDF-Format.', icon: 'file-check' },
    ],
    faq: [
      { question: 'Was ist WebP?', answer: 'Ein von Google entwickeltes Format für hohe Kompression im Web.' },
      { question: 'Ist die Konvertierung verlustfrei?', answer: 'Ja, die Bildqualität des WebP-Originals bleibt im PDF erhalten.' },
      { question: 'Funktioniert es mit animierten WebPs?', answer: 'Diese werden als Standbilder konvertiert.' },
    ],
  },

  'svg-to-pdf': {
    title: 'SVG in PDF',
    metaDescription: 'SVG-Vektorgrafiken in PDF konvertieren. Erhalten Sie Skalierbarkeit und verlustfreie Qualität.',
    keywords: ['svg in pdf', 'vektorgrafik zu pdf', 'skalierbares pdf', 'logo konvertieren'],
    description: `
      <p>Wandeln Sie skalierbare Vektorgrafiken (SVG) in PDF um, ohne an Schärfe zu verlieren. Da PDF ebenfalls Vektoren unterstützt, bleibt Ihr Design bei jeder Vergrößerung gestochen scharf.</p>
    `,
    howToUse: [
      { step: 1, title: 'SVG hochladen', description: 'Wählen Sie Ihre SVG-Vektordateien aus.' },
      { step: 2, title: 'Seiteneinstellungen', description: 'Wählen Sie ein passendes Format für Ihre Grafiken.' },
      { step: 3, title: 'Konvertieren', description: 'Laden Sie das Vektor-PDF herunter.' },
    ],
    useCases: [
      { title: 'Logo-Druck', description: 'Bereiten Sie Vektor-Logos für den professionellen Druck vor.', icon: 'award' },
      { title: 'Technische Pläne', description: 'Wandeln Sie CAD-Exporte oder Diagramme in PDF um.', icon: 'ruler' },
      { title: 'Illustrationen', description: 'Erstellen Sie hochauflösende Dokumente aus Vektor-Kunstwerken.', icon: 'grid' },
    ],
    faq: [
      { question: 'Bleibt die Vektorqualität erhalten?', answer: 'Ja, das PDF behält die mathematischen Pfade bei, sodass der Inhalt unendlich skalierbar bleibt.' },
      { question: 'Werden Schriften korrekt eingebettet?', answer: 'In der Regel ja. Für beste Ergebnisse sollten Schriften im SVG in Pfade umgewandelt sein.' },
      { question: 'Unterstützt das Tool komplexe Filter?', answer: 'Die meisten Standard-SVG-Filter und Gradienten werden unterstützt.' },
    ],
  },

  'heic-to-pdf': {
    title: 'HEIC in PDF',
    metaDescription: 'iPhone HEIC-Fotos in PDF konvertieren. Apple-Bildformate einfach für Windows und Android nutzbar machen.',
    keywords: ['heic in pdf', 'iphone foto konvertieren', 'apple bild zu pdf', 'heic zu pdf deutsch'],
    description: `
      <p>Konvertieren Sie HEIC-Fotos von Ihrem iPhone oder iPad direkt in PDF. HEIC bietet zwar eine gute Kompression, ist aber nicht überall lesbar – PDF hingegen schon.</p>
    `,
    howToUse: [
      { step: 1, title: 'HEIC hochladen', description: 'Wählen Sie Ihre Apple-Fotos aus.' },
      { step: 2, title: 'Sortieren', description: 'Bringen Sie Ihre Fotos in die gewünschte Reihenfolge.' },
      { step: 3, title: 'Download', description: 'Laden Sie das universelle PDF-Dokument herunter.' },
    ],
    useCases: [
      { title: 'Mobile Fotoalben', description: 'Erstellen Sie PDFs aus Ihren Handyfotos zum Teilen mit Nicht-Apple-Nutzern.', icon: 'smartphone' },
      { title: 'Scans vom iPhone', description: 'Wandeln Sie mit der Kamera aufgenommene Dokumente in saubere PDFs um.', icon: 'scan' },
      { title: 'Kompatibilität', description: 'Machen Sie HEIC-Bilder für Windows-PCs und Android-Geräte nutzbar.', icon: 'share-2' },
    ],
    faq: [
      { question: 'Was ist HEIC?', answer: 'Das Standard-Bildformat von Apple für effiziente Speicherung.' },
      { question: 'Werden Live-Photos unterstützt?', answer: 'Das Tool konvertiert das Hauptbild des Live-Photos in das PDF.' },
      { question: 'Bleiben Metadaten (Exif) erhalten?', answer: 'Sie können wählen, ob Aufnahmeort und Datum im Dokument verbleiben sollen.' },
    ],
  },

  'txt-to-pdf': {
    title: 'Text in PDF',
    metaDescription: 'Konvertieren Sie Textdateien (TXT) in formatiertes PDF. Passen Sie Schriftart und Layout an.',
    keywords: ['txt in pdf', 'textdatei umwandeln', 'quellcode zu pdf', 'notizen zu pdf'],
    description: `
      <p>Wandeln Sie einfache Textdateien (.txt) in formatierte PDF-Dokumente um. Passen Sie Schriftarten, Ränder und das Seitenlayout an, um aus simplen Notizen professionelle Dokumente zu machen.</p>
    `,
    howToUse: [
      { step: 1, title: 'Textdatei hochladen', description: 'Wählen Sie Ihre .txt-Datei aus.' },
      { step: 2, title: 'Formatierung', description: 'Wählen Sie Schriftart (z. B. Monospace für Code) und Seitenränder.' },
      { step: 3, title: 'Speichern', description: 'Laden Sie das formatierte PDF-Dokument herunter.' },
    ],
    useCases: [
      { title: 'Code-Dokumentation', description: 'Wandeln Sie Programmiercode in lesbare PDF-Dateien um.', icon: 'code' },
      { title: 'Logfile-Archivierung', description: 'Speichern Sie Server-Logs in einem festen Dokumentenformat.', icon: 'file-text' },
      { title: 'Manuskripte', description: 'Machen Sie aus einfachen Textentwürfen druckfertige PDFs.', icon: 'sticky-note' },
    ],
    faq: [
      { question: 'Werden Sonderzeichen unterstützt?', answer: 'Ja, das Tool unterstützt UTF-8 Kodierung für internationale Schriftzeichen.' },
      { question: 'Gibt es einen automatischen Zeilenumbruch?', answer: 'Ja, zu lange Zeilen werden automatisch an die Seitenbreite angepasst.' },
      { question: 'Kann ich die Schriftgröße ändern?', answer: 'Ja, Sie können die Größe für optimale Lesbarkeit einstellen.' },
    ],
  },

  'json-to-pdf': {
    title: 'JSON in PDF',
    metaDescription: 'Konvertieren Sie JSON-Dateien in formatiertes PDF. Mit Syntax-Highlighting und strukturierter Ausgabe.',
    keywords: ['json in pdf', 'json visualisieren', 'api daten zu pdf', 'json formatieren'],
    description: `
      <p>Verwandeln Sie JSON-Daten in ein lesbares, schön formatiertes PDF. Das Tool bietet automatisches Syntax-Highlighting und Einrückungen, um komplexe Datenstrukturen visuell aufzubereiten.</p>
    `,
    howToUse: [
      { step: 1, title: 'JSON hochladen', description: 'Wählen Sie Ihre .json-Datei aus.' },
      { step: 2, title: 'Stil wählen', description: 'Konfigurieren Sie das Farbschema für das Syntax-Highlighting.' },
      { step: 3, title: 'Generieren', description: 'Laden Sie das strukturierte Daten-PDF herunter.' },
    ],
    useCases: [
      { title: 'API-Dokumentation', description: 'Präsentieren Sie Beispiel-Responses in einem sauberen Dokument.', icon: 'code' },
      { title: 'Konfigurations-Backups', description: 'Archivieren Sie Einstellungen in einem menschenlesbaren Format.', icon: 'settings' },
      { title: 'Daten-Berichte', description: 'Erstellen Sie Berichte aus JSON-Datenexporten.', icon: 'bar-chart' },
    ],
    faq: [
      { question: 'Ist der Code farbig markiert?', answer: 'Ja, Schlüssel, Werte und Datentypen werden zur besseren Übersicht farblich hervorgehoben.' },
      { question: 'Wie werden große Dateien gehandhabt?', answer: 'Lange JSON-Strukturen werden automatisch auf mehrere Seiten verteilt.' },
      { question: 'Brauche ich Programmierkenntnisse?', answer: 'Nein, laden Sie einfach die Datei hoch und das Tool übernimmt das Layout.' },
    ],
  },// ==================== CONVERT FROM PDF ====================
  'pdf-to-jpg': {
    title: 'PDF in JPG',
    metaDescription: 'Konvertieren Sie PDF-Seiten in JPG-Bilder. Hochwertige Extraktion mit einstellbarer Auflösung.',
    keywords: ['pdf in jpg', 'pdf zu jpeg', 'pdf als bild speichern', 'seiten extrahieren'],
    description: `
      <p>Wandeln Sie Ihre PDF-Seiten in hochwertige JPG-Bilder um. Sie können entweder alle Seiten konvertieren oder gezielt einzelne Seiten auswählen. Dabei lassen sich Auflösung (DPI) und Bildqualität individuell anpassen.</p>
      <p>Ideal zum Erstellen von Vorschaubildern, zum Teilen von Dokumenten in sozialen Netzwerken oder zur Verwendung in Bildbearbeitungsprogrammen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie die PDF-Datei aus, die Sie umwandeln möchten.' },
      { step: 2, title: 'Qualität wählen', description: 'Stellen Sie die gewünschte Auflösung (DPI) und den Kompressionsgrad ein.' },
      { step: 3, title: 'Konvertieren', description: 'Laden Sie die Bilder einzeln oder gesammelt als ZIP-Archiv herunter.' },
    ],
    useCases: [
      { title: 'Web-Veröffentlichung', description: 'Erstellen Sie Bildversionen von PDF-Seiten für Ihre Website.', icon: 'globe' },
      { title: 'Social Media', description: 'Teilen Sie Dokumentinhalte als einfache Bilder auf Instagram oder LinkedIn.', icon: 'share-2' },
      { title: 'Präsentationen', description: 'Fügen Sie PDF-Folien als Bilder in PowerPoint oder Keynote ein.', icon: 'presentation' },
    ],
    faq: [
      { question: 'Welche Auflösung wird unterstützt?', answer: 'Sie können zwischen 72 DPI (Web) und 300 DPI (Druckqualität) wählen.' },
      { question: 'Kann ich nur bestimmte Seiten wählen?', answer: 'Ja, Sie können einzelne Seiten oder Seitenbereiche für die Konvertierung festlegen.' },
      { question: 'Wie erhalte ich die Dateien?', answer: 'Bei mehreren Seiten erstellt das Tool automatisch einen praktischen ZIP-Ordner.' },
    ],
  },

  'pdf-to-png': {
    title: 'PDF in PNG',
    metaDescription: 'Konvertieren Sie PDF-Seiten in PNG-Bilder. Verlustfreie Qualität mit Unterstützung für Transparenz.',
    keywords: ['pdf in png', 'pdf zu bild verlustfrei', 'grafik extraktion pdf', 'transparenz pdf'],
    description: `
      <p>Konvertieren Sie PDF-Dokumente in das verlustfreie PNG-Format. Im Gegensatz zu JPG bietet PNG eine perfekte Bildqualität ohne Artefakte und unterstützt transparente Hintergründe.</p>
      <p>Besonders geeignet für PDFs, die Diagramme, Logos oder Texte enthalten, die gestochen scharf bleiben müssen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Laden Sie Ihr Dokument hoch.' },
      { step: 2, title: 'Optionen konfigurieren', description: 'Wählen Sie die Seiten und die gewünschte Pixeldichte aus.' },
      { step: 3, title: 'PNGs speichern', description: 'Extrahieren Sie die Seiten als hochwertige PNG-Dateien.' },
    ],
    useCases: [
      { title: 'Grafik-Extraktion', description: 'Speichern Sie Vektorgrafiken aus PDFs als saubere Rasterbilder.', icon: 'image' },
      { title: 'Design-Assets', description: 'Wandeln Sie PDF-Entwürfe in PNGs für Grafiksoftware um.', icon: 'palette' },
      { title: 'Technische Doku', description: 'Erstellen Sie scharfe Abbildungen für Handbücher und Anleitungen.', icon: 'file-text' },
    ],
    faq: [
      { question: 'Warum PNG statt JPG?', answer: 'PNG ist verlustfrei und eignet sich besser für Texte und Grafiken mit harten Kanten.' },
      { question: 'Bleibt die Transparenz erhalten?', answer: 'Ja, sofern das PDF transparente Ebenen hat, werden diese im PNG übernommen.' },
      { question: 'Welchen DPI-Wert soll ich nutzen?', answer: '150 DPI für den Bildschirm, 300 DPI für eine sehr hohe Detailschärfe.' },
    ],
  },

  'pdf-to-webp': {
    title: 'PDF in WebP',
    metaDescription: 'Konvertieren Sie PDF-Seiten in WebP-Bilder. Modernes Format mit exzellenter Kompression für das Web.',
    keywords: ['pdf in webp', 'modernes bildformat', 'web optimierte bilder', 'pdf konverter'],
    description: `
      <p>Wandeln Sie PDF-Seiten in das moderne WebP-Format von Google um. WebP bietet eine deutlich bessere Kompression als JPG oder PNG bei vergleichbarer Qualität.</p>
      <p>Dies ist die beste Wahl, wenn Sie PDF-Inhalte schnell ladend auf einer modernen Website anzeigen möchten.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das PDF-Dokument aus.' },
      { step: 2, title: 'Kompression wählen', description: 'Stellen Sie die Balance zwischen Dateigröße und Qualität ein.' },
      { step: 3, title: 'Download', description: 'Laden Sie die web-optimierten WebP-Bilder herunter.' },
    ],
    useCases: [
      { title: 'Web-Optimierung', description: 'Reduzieren Sie die Ladezeiten Ihrer Website durch WebP-Bilder.', icon: 'globe' },
      { title: 'Bandbreite sparen', description: 'Ideal für mobile Anwendungen mit begrenztem Datenvolumen.', icon: 'zap' },
      { title: 'Modernes Web-Design', description: 'Nutzen Sie zukunftssichere Bildformate für Ihre Projekte.', icon: 'layout' },
    ],
    faq: [
      { question: 'Ist WebP mit allen Browsern kompatibel?', answer: 'Ja, alle modernen Browser wie Chrome, Firefox, Edge und Safari unterstützen WebP.' },
      { question: 'Wie viel kleiner sind die Dateien?', answer: 'WebP-Dateien sind oft 25-35 % kleiner als vergleichbare JPG-Dateien.' },
      { question: 'Gibt es Qualitätsverluste?', answer: 'WebP bietet sowohl verlustbehaftete als auch verlustfreie Kompression an.' },
    ],
  },

  'pdf-to-bmp': {
    title: 'PDF in BMP',
    metaDescription: 'PDF-Seiten in BMP-Bitmap-Bilder konvertieren. Unkomprimiertes Format für maximale Kompatibilität.',
    keywords: ['pdf in bmp', 'bitmap konverter', 'unkomprimierte bilder', 'legacy format'],
    description: `
      <p>Wandeln Sie PDF-Seiten in das klassische BMP-Format (Windows Bitmap) um. BMP ist ein unkomprimiertes Format, das eine universelle Kompatibilität mit älteren Systemen und speziellen Windows-Anwendungen garantiert.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie Ihre Datei aus.' },
      { step: 2, title: 'Seiten wählen', description: 'Bestimmen Sie, welche Seiten als Bitmap gespeichert werden sollen.' },
      { step: 3, title: 'BMP erstellen', description: 'Konvertieren und laden Sie die Bitmap-Bilder herunter.' },
    ],
    useCases: [
      { title: 'Altsysteme', description: 'Erstellen Sie Bilder für Software, die keine modernen Formate unterstützt.', icon: 'history' },
      { title: 'Windows-Anwendungen', description: 'Generieren Sie kompatible Dateien für spezifische Windows-Tools.', icon: 'monitor' },
      { title: 'Verlustfreie Archivierung', description: 'Speichern Sie Bilder völlig ohne Kompressionsartefakte.', icon: 'archive' },
    ],
    faq: [
      { question: 'Warum sollte ich heute noch BMP nutzen?', answer: 'Hauptsächlich für die Kompatibilität mit Legacy-Software oder industriellen Anwendungen.' },
      { question: 'Sind BMP-Dateien sehr groß?', answer: 'Ja, da sie nicht komprimiert sind, sind sie deutlich größer als JPG oder PNG.' },
      { question: 'Welche Farbtiefe wird unterstützt?', answer: 'Das Tool unterstützt Standard-Bitmaps mit 24-Bit und 32-Bit.' },
    ],
  },

  'pdf-to-tiff': {
    title: 'PDF in TIFF',
    metaDescription: 'PDF in TIFF-Bilder konvertieren. Professionelle Qualität mit Unterstützung für mehrseitige TIFF-Dateien.',
    keywords: ['pdf in tiff', 'profidruck tiff', 'mehrseitiges tiff', 'archivierung'],
    description: `
      <p>Konvertieren Sie PDFs in das hochwertige TIFF-Format. TIFF ist der Standard im professionellen Druck und in der Langzeitarchivierung, da es eine extrem hohe Farbtiefe und verlustfreie Kompression (LZW/ZIP) unterstützt.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Laden Sie das Dokument hoch.' },
      { step: 2, title: 'Format wählen', description: 'Wählen Sie zwischen einzelnen TIFFs pro Seite oder einem einzelnen, mehrseitigen TIFF.' },
      { step: 3, title: 'Download', description: 'Laden Sie die professionellen Bilddateien herunter.' },
    ],
    useCases: [
      { title: 'Professioneller Druck', description: 'Erstellen Sie druckfähige TIFF-Dateien für Verlage und Druckereien.', icon: 'printer' },
      { title: 'Dokumenten-Archivierung', description: 'Sichern Sie Dokumente in einem stabilen, hochwertigen Archivformat.', icon: 'archive' },
      { title: 'Publikationen', description: 'Wandeln Sie PDFs für die Weiterverarbeitung in Satzprogrammen um.', icon: 'book' },
    ],
    faq: [
      { question: 'Unterstützt das Tool mehrseitige TIFFs?', answer: 'Ja, Sie können das gesamte PDF in eine einzige, mehrseitige TIFF-Datei umwandeln.' },
      { question: 'Welche Kompression wird genutzt?', answer: 'Sie können zwischen LZW, ZIP oder völlig unkomprimierter Ausgabe wählen.' },
      { question: 'Welcher DPI-Wert ist für den Druck nötig?', answer: 'Für professionelle Ergebnisse empfehlen wir mindestens 300 DPI.' },
    ],
  },

  'pdf-to-greyscale': {
    title: 'PDF in Graustufen',
    metaDescription: 'Farb-PDFs in Schwarz-Weiß (Graustufen) konvertieren. Dateigröße reduzieren und Druckkosten sparen.',
    keywords: ['pdf graustufen', 'pdf schwarz weiß machen', 'tinte sparen pdf', 'farbe entfernen'],
    description: `
      <p>Wandeln Sie ein farbiges PDF in eine Graustufen-Version (Schwarz-Weiß) um. Dies ist ideal, um die Dateigröße zu reduzieren und Dokumente optimal für den Schwarz-Weiß-Druck vorzubereiten.</p>
      <p>Texte bleiben gestochen scharf und Bilder behalten ihre Details, während alle Farbinformationen entfernt werden.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das farbige Dokument aus.' },
      { step: 2, title: 'Vorschau prüfen', description: 'Sehen Sie sich an, wie die Graustufen-Konvertierung wirkt.' },
      { step: 3, title: 'Download', description: 'Laden Sie das optimierte Schwarz-Weiß-PDF herunter.' },
    ],
    useCases: [
      { title: 'Druckkosten sparen', description: 'Vermeiden Sie teuren Farbdruck bei Entwürfen oder Skripten.', icon: 'printer' },
      { title: 'Dateigröße verringern', description: 'Reduzieren Sie die Dateigröße durch das Entfernen von Farbkanälen.', icon: 'minimize-2' },
      { title: 'Professionelle Ästhetik', description: 'Geben Sie Dokumenten einen klassischen Schwarz-Weiß-Look.', icon: 'file-text' },
    ],
    faq: [
      { question: 'Bleibt der Text lesbar?', answer: 'Absolut. Das Tool optimiert die Kontraste, damit Texte weiterhin perfekt lesbar sind.' },
      { question: 'Wie viel Speicherplatz spare ich?', answer: 'Bei bildreichen Dokumenten kann die Größe oft um 20-50 % sinken.' },
      { question: 'Kann ich nur bestimmte Seiten entfärben?', answer: 'Ja, Sie können gezielt auswählen, welche Seiten konvertiert werden sollen.' },
    ],
  },

  'pdf-to-json': {
    title: 'PDF in JSON',
    metaDescription: 'Extrahiert PDF-Inhalte in das JSON-Format. Erhalten Sie strukturierte Daten für Ihre Anwendungen.',
    keywords: ['pdf in json', 'datenextraktion pdf', 'pdf parser', 'strukturierte daten'],
    description: `
      <p>Extrahieren Sie Text, Metadaten und die Struktur Ihres PDF-Dokuments in das maschinenlesbare JSON-Format. Dies ist das perfekte Tool für Entwickler und Datenanalysten.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie die PDF-Quelldatei aus.' },
      { step: 2, title: 'Daten wählen', description: 'Bestimmen Sie, welche Informationen (Text, Layout, Metadaten) extrahiert werden sollen.' },
      { step: 3, title: 'JSON exportieren', description: 'Laden Sie die fertige JSON-Datei für Ihre Programmierung herunter.' },
    ],
    useCases: [
      { title: 'Datenanalyse', description: 'Werten Sie Textinhalte aus PDFs automatisiert aus.', icon: 'database' },
      { title: 'System-Integration', description: 'Importieren Sie PDF-Inhalte direkt in Ihre Datenbank oder App.', icon: 'plug' },
      { title: 'Metadaten-Audit', description: 'Analysieren Sie die technischen Details einer großen Anzahl von PDFs.', icon: 'search' },
    ],
    faq: [
      { question: 'Was genau wird im JSON gespeichert?', answer: 'Textinhalte, Positionsdaten, Schriftarten, Seitendimensionen und Metadaten.' },
      { question: 'Funktioniert das bei Scans?', answer: 'Nur wenn diese zuvor mit unserem OCR-Tool durchsuchbar gemacht wurden.' },
      { question: 'Ist das Format dokumentiert?', answer: 'Ja, wir nutzen ein standardisiertes Schema für eine einfache Weiterverarbeitung.' },
    ],
  },// ==================== ORGANIZE & MANAGE ====================
  'alternate-merge': {
    title: 'Wechselweise zusammenfügen',
    metaDescription: 'PDFs durch abwechselnde Seiten zusammenfügen. Kombinieren Sie Vorder- und Rückseiten-Scans in einem Dokument.',
    keywords: ['alternate merge', 'pdf verschachteln', 'vorderseite rückseite kombinieren', 'scans zusammenfügen'],
    description: `
      <p>Mit "Wechselweise zusammenfügen" kombinieren Sie zwei PDF-Dateien, indem die Seiten abwechselnd (interleaved) hintereinandergelegt werden. Dies ist die perfekte Lösung, wenn Sie Vorder- und Rückseiten separat gescannt haben.</p>
      <p>Laden Sie einfach die Datei mit den Vorderseiten und die Datei mit den Rückseiten hoch. Das Tool fügt diese automatisch zu einem logischen Gesamtdokument zusammen. Sie können die Reihenfolge einer Datei auch umkehren, falls diese rückwärts gescannt wurde.</p>
    `,
    howToUse: [
      { step: 1, title: 'Zwei PDFs hochladen', description: 'Wählen Sie die PDF mit den ungeraden Seiten (Vorderseiten) und die mit den geraden Seiten (Rückseiten).' },
      { step: 2, title: 'Reihenfolge konfigurieren', description: 'Wählen Sie bei Bedarf "Reihenfolge umkehren" für das zweite Dokument (oft bei Back-to-Front-Scans nötig).' },
      { step: 3, title: 'Zusammenfügen', description: 'Klicken Sie auf "Merge", um die Seiten zu verschachteln und das Ergebnis zu laden.' },
    ],
    useCases: [
      { title: 'Duplex-Scannen', description: 'Kombinieren Sie Scans von Geräten, die nicht automatisch beidseitig scannen.', icon: 'copy' },
      { title: 'Dokumentenmontage', description: 'Verschachteln Sie Seiten aus zwei zusammengehörigen Berichten.', icon: 'layers' },
      { title: 'Buch-Scans', description: 'Fügen Sie Scans von linken und rechten Buchseiten zu einem flüssigen Dokument zusammen.', icon: 'book' },
    ],
    faq: [
      { question: 'Was passiert bei unterschiedlicher Seitenanzahl?', answer: 'Überschüssige Seiten des längeren Dokuments werden einfach am Ende angehängt.' },
      { question: 'Kann ich die Seitenreihenfolge umkehren?', answer: 'Ja, das Tool bietet eine Option, um eine der Dateien vor dem Mischen umzukehren.' },
      { question: 'Ist das besser als normales Zusammenfügen?', answer: 'Ja, beim normalen Mischen werden Dateien nur angehängt; hier werden sie wie ein Kartenspiel gemischt.' },
    ],
  },

  'add-attachments': {
    title: 'Anhänge hinzufügen',
    metaDescription: 'Dateien in PDF-Dokumente einbetten. Fügen Sie beliebige Dateitypen als Anhang zu Ihren PDFs hinzu.',
    keywords: ['pdf anhänge', 'dateien in pdf einbetten', 'pdf portfolio', 'datei an pdf hängen'],
    description: `
      <p>Betten Sie beliebige Dateien direkt in Ihre PDF-Dokumente ein. Ob Tabellen, Bilder oder Quellcode – erstellen Sie umfassende PDF-Pakete, in denen alle relevanten Daten enthalten sind.</p>
      <p>Die Anhänge werden Teil der PDF-Datei und können vom Empfänger mit jedem gängigen PDF-Reader wieder extrahiert werden.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das Haupt-PDF-Dokument aus.' },
      { step: 2, title: 'Dateien hinzufügen', description: 'Wählen Sie die Dateien aus, die Sie im PDF einbetten möchten.' },
      { step: 3, title: 'Speichern', description: 'Laden Sie das PDF mit den integrierten Anhängen herunter.' },
    ],
    useCases: [
      { title: 'Projektpakete', description: 'Bündeln Sie Designdateien oder Kalkulationen direkt in der Dokumentation.', icon: 'package' },
      { title: 'Berichtsverteilung', description: 'Fügen Sie Rohdaten als Excel-Datei an einen PDF-Bericht an.', icon: 'paperclip' },
      { title: 'Vertragsunterlagen', description: 'Hängen Sie unterstützende Dokumente direkt an den Hauptvertrag.', icon: 'file-text' },
    ],
    faq: [
      { question: 'Welche Dateitypen werden unterstützt?', answer: 'Sie können jeden beliebigen Dateityp in ein PDF einbetten.' },
      { question: 'Gibt es eine Größenbeschränkung?', answer: 'Die Gesamtgröße des PDFs inkl. Anhängen sollte 500 MB nicht überschreiten.' },
      { question: 'Können Empfänger die Dateien sehen?', answer: 'Ja, moderne PDF-Reader zeigen Anhänge in einer speziellen Seitenleiste an.' },
    ],
  },

  'extract-attachments': {
    title: 'Anhänge extrahieren',
    metaDescription: 'Eingebettete Dateien aus PDFs extrahieren. Laden Sie alle Anhänge aus einem PDF-Dokument herunter.',
    keywords: ['anhänge extrahieren', 'pdf dateien extrahieren', 'eingebettete dateien laden'],
    description: `
      <p>Holen Sie alle eingebetteten Dateien aus einem PDF-Dokument heraus. Sie können Anhänge einzeln oder alle zusammen als praktisches ZIP-Archiv herunterladen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das PDF aus, das Anhänge enthält.' },
      { step: 2, title: 'Anhänge anzeigen', description: 'Sehen Sie eine Liste aller im PDF versteckten Dateien.' },
      { step: 3, title: 'Herunterladen', description: 'Speichern Sie die Dateien einzeln oder als komplettes Paket.' },
    ],
    useCases: [
      { title: 'Quellcode abrufen', description: 'Extrahierten Sie Originaldaten aus wissenschaftlichen PDF-Berichten.', icon: 'download' },
      { title: 'Pakete entpacken', description: 'Greifen Sie auf Dokumente zu, die Ihnen als PDF-Portfolio geschickt wurden.', icon: 'folder-open' },
      { title: 'Batch-Extraktion', description: 'Holen Sie Anhänge aus mehreren PDFs gleichzeitig heraus.', icon: 'layers' },
    ],
    faq: [
      { question: 'Was, wenn keine Anhänge vorhanden sind?', answer: 'Das Tool zeigt Ihnen sofort an, ob das Dokument eingebettete Dateien enthält oder nicht.' },
      { question: 'Wird das PDF dabei beschädigt?', answer: 'Nein, die Anhänge werden kopiert; das Original-PDF bleibt unverändert.' },
      { question: 'Werden alle Formate unterstützt?', answer: 'Ja, das Tool extrahiert jeden Dateityp, der im PDF gespeichert wurde.' },
    ],
  },

  'divide-pages': {
    title: 'Seiten unterteilen',
    metaDescription: 'PDF-Seiten in mehrere Abschnitte aufteilen. Unterteilen Sie Seiten horizontal oder vertikal.',
    keywords: ['pdf seiten teilen', 'seite schneiden', 'pdf sektionieren', 'rasterteilung'],
    description: `
      <p>Dieses Tool schneidet einzelne PDF-Seiten in mehrere Teile. Sie können Seiten horizontal, vertikal oder in einem Raster unterteilen, um aus einer Seite mehrere neue Seiten zu machen.</p>
      <p>Besonders nützlich für Scans, bei denen mehrere Dokumente (z. B. Quittungen) auf einer Seite liegen, oder um großformatige Pläne handlich zu machen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Laden Sie das PDF-Dokument hoch.' },
      { step: 2, title: 'Teilung festlegen', description: 'Wählen Sie zwischen horizontaler/vertikaler Teilung oder einem Gitter (z. B. 2x2).' },
      { step: 3, title: 'Schneiden', description: 'Laden Sie das PDF mit den nun unterteilten Einzelseiten herunter.' },
    ],
    useCases: [
      { title: 'Scans trennen', description: 'Teilen Sie eine Seite mit mehreren eingescannten Belegen in Einzelbilder.', icon: 'scissors' },
      { title: 'Großformate anpassen', description: 'Schneiden Sie ein A3-Dokument in zwei A4-Seiten.', icon: 'maximize-2' },
      { title: 'Karten erstellen', description: 'Unterteilen Sie Seiten in visitenkartengroße Abschnitte.', icon: 'grid' },
    ],
    faq: [
      { question: 'Sind die Schnitte präzise?', answer: 'Ja, das Tool teilt die Seite exakt nach Ihren Vorgaben auf.' },
      { question: 'Was passiert mit Text auf der Schnittlinie?', answer: 'Inhalte direkt auf der Linie werden getrennt; achten Sie auf ausreichende Abstände.' },
      { question: 'Kann ich nur bestimmte Seiten teilen?', answer: 'Ja, Sie können die Teilung auf das gesamte Dokument oder Seitenauswahlen anwenden.' },
    ],
  },

  'n-up-pdf': {
    title: 'N-Up (Mehrere Seiten pro Blatt)',
    metaDescription: 'Drucken Sie mehrere PDF-Seiten auf ein einzelnes Blatt. Erstellen Sie 2-Up, 4-Up oder individuelle Layouts.',
    keywords: ['n-up pdf', 'mehrere seiten pro blatt', 'pdf impositon', 'handout erstellen'],
    description: `
      <p>N-Up PDF ordnet mehrere Seiten Ihres Dokuments auf einem einzigen Blatt an (z. B. 2, 4, 6 oder 9 Seiten pro Blatt). Dies spart Papier beim Drucken und ist ideal für Handouts.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das zu druckende Dokument.' },
      { step: 2, title: 'Layout wählen', description: 'Entscheiden Sie sich für ein Raster (z. B. 2x2 für 4 Seiten pro Blatt).' },
      { step: 3, title: 'Generieren', description: 'Laden Sie das optimierte Layout herunter.' },
    ],
    useCases: [
      { title: 'Papier sparen', description: 'Reduzieren Sie den Papierverbrauch beim Korrekturlesen.', icon: 'leaf' },
      { title: 'Präsentations-Handouts', description: 'Erstellen Sie kompakte Übersichten Ihrer Folien.', icon: 'file-text' },
      { title: 'Broschüren-Vorbereitung', description: 'Ordnen Sie Seiten für spezielle Druckformate an.', icon: 'eye' },
    ],
    faq: [
      { question: 'Bleibt die Reihenfolge erhalten?', answer: 'Ja, die Seiten werden standardmäßig von links nach rechts und oben nach unten angeordnet.' },
      { question: 'Kann ich Ränder hinzufügen?', answer: 'Ja, Sie können Abstände zwischen den einzelnen Kacheln festlegen.' },
      { question: 'Wird die Qualität schlechter?', answer: 'Nein, die Seiten werden lediglich verkleinert, die Details bleiben scharf.' },
    ],
  },

  'reverse-pages': {
    title: 'Seitenreihenfolge umkehren',
    metaDescription: 'Kehren Sie die Reihenfolge der PDF-Seiten um. Spiegeln Sie das Dokument von hinten nach vorne.',
    keywords: ['pdf umkehren', 'seiten spiegeln', 'rückwärts sortieren', 'scan korrigieren'],
    description: `
      <p>Dieses Tool dreht die gesamte Seitenreihenfolge Ihres PDFs um. Die letzte Seite wird zur ersten, die vorletzte zur zweiten usw. Ideal für Dokumente, die falsch herum eingescannt wurden.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie die Datei aus.' },
      { step: 2, title: 'Bereich wählen', description: 'Wählen Sie das gesamte Dokument oder nur einen Teilbereich zum Umkehren.' },
      { step: 3, title: 'Umkehren', description: 'Speichern Sie das PDF mit der neuen Sortierung.' },
    ],
    useCases: [
      { title: 'Scan-Fehler beheben', description: 'Korrigieren Sie Stapel-Scans, die in der falschen Reihenfolge eingelesen wurden.', icon: 'refresh-cw' },
      { title: 'Druckvorbereitung', description: 'Bereiten Sie Dokumente für Drucker vor, die Seiten in umgekehrter Folge ausgeben.', icon: 'printer' },
      { title: 'Prüfprozesse', description: 'Betrachten Sie Dokumente schnell aus einer anderen Perspektive.', icon: 'arrow-up-down' },
    ],
    faq: [
      { question: 'Werden Lesezeichen angepasst?', answer: 'Ja, die interne Verlinkung der Lesezeichen wird automatisch auf die neuen Seitenpositionen aktualisiert.' },
      { question: 'Ist das das Gleiche wie Drehen?', answer: 'Nein, Drehen ändert die Ausrichtung der Seite, Umkehren ändert die Position im Dokument.' },
      { question: 'Kann ich nur das Ende umdrehen?', answer: 'Ja, Sie können einen spezifischen Seitenbereich (z. B. Seite 10-20) definieren.' },
    ],
  },

  'compare-pdfs': {
    title: 'PDFs vergleichen',
    metaDescription: 'Vergleichen Sie zwei PDF-Dokumente. Lassen Sie sich Unterschiede zwischen Versionen farblich hervorheben.',
    keywords: ['pdf vergleichen', 'dokumenten-diff', 'versionskontrolle pdf', 'änderungen finden'],
    description: `
      <p>Analysieren Sie zwei Versionen eines Dokuments auf einen Blick. Das Tool markiert Textänderungen, Hinzufügungen und Löschungen farblich, sodass Sie Revisionen sofort prüfen können.</p>
    `,
    howToUse: [
      { step: 1, title: 'Zwei PDFs hochladen', description: 'Laden Sie das Original und die bearbeitete Version hoch.' },
      { step: 2, title: 'Vergleich starten', description: 'Das Tool analysiert beide Dateien auf Text- und Layout-Differenzen.' },
      { step: 3, title: 'Ergebnisse prüfen', description: 'Betrachten Sie die Highlights direkt im Browser oder laden Sie einen Bericht herunter.' },
    ],
    useCases: [
      { title: 'Vertragsprüfung', description: 'Finden Sie versteckte Änderungen in neuen Vertragsentwürfen.', icon: 'file-text' },
      { title: 'Lektorat', description: 'Kontrollieren Sie, ob alle Korrekturwünsche korrekt umgesetzt wurden.', icon: 'git-compare' },
      { title: 'Qualitätssicherung', description: 'Stellen Sie sicher, dass beim Konvertieren keine Inhalte verloren gingen.', icon: 'check-circle' },
    ],
    faq: [
      { question: 'Was wird erkannt?', answer: 'Textänderungen, gelöschte Abschnitte und oft auch verschobene Bilder.' },
      { question: 'Geht das auch bei Scans?', answer: 'Dafür müssen die Scans zuerst per OCR lesbar gemacht werden.' },
      { question: 'Wie werden Unterschiede gezeigt?', answer: 'Meist durch farbige Unterlegungen (Rot für gelöscht, Grün für neu).' },
    ],
  },// ==================== OPTIMIZE & REPAIR ====================
  'fix-page-size': {
    title: 'Seitengröße vereinheitlichen',
    metaDescription: 'Standardisieren Sie PDF-Seitengrößen. Konvertieren Sie alle Seiten in ein einheitliches Format.',
    keywords: ['pdf seitengröße anpassen', 'pdf vereinheitlichen', 'seitenformat korrigieren', 'pdf auf a4'],
    description: `
      <p>Bringen Sie alle Seiten Ihres PDFs auf ein einheitliches Format. Dieses Tool konvertiert Dokumente mit gemischten Seitengrößen in ein konsistentes Format für eine professionelle Präsentation oder den Druck.</p>
      <p>Wählen Sie aus Standardgrößen wie A4 oder US-Letter oder definieren Sie eigene Maße. Der Inhalt wird dabei passgenau skaliert oder zentriert.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie das PDF mit unterschiedlichen Formaten aus.' },
      { step: 2, title: 'Zielformat wählen', description: 'Wählen Sie A4, Letter oder geben Sie eigene Maße in mm/Zoll ein.' },
      { step: 3, title: 'Anwenden', description: 'Laden Sie das PDF mit den nun einheitlichen Seitenmaßen herunter.' },
    ],
    useCases: [
      { title: 'Druckvorbereitung', description: 'Stellen Sie sicher, dass alle Seiten ohne Skalierungsfehler gedruckt werden.', icon: 'printer' },
      { title: 'Dokumenten-Cleanup', description: 'Korrigieren Sie unschöne Format-Mixe nach dem Zusammenfügen verschiedener Dateien.', icon: 'file-check' },
      { title: 'Professionelle Dossiers', description: 'Erstellen Sie einheitliche Unterlagen für Kunden oder Behörden.', icon: 'briefcase' },
    ],
    faq: [
      { question: 'Wird der Inhalt verzerrt?', answer: 'Nein, Sie können wählen, ob der Inhalt proportional skaliert oder lediglich zentriert werden soll.' },
      { question: 'Kann ich das Seitenverhältnis beibehalten?', answer: 'Ja, das Tool bietet Optionen zur intelligenten Anpassung ohne Verzerrung.' },
      { question: 'Welche Größen sind verfügbar?', answer: 'Alle gängigen ISO-Formate (A0-A5), US-Formate und benutzerdefinierte Maße.' },
    ],
  },

  'linearize-pdf': {
    title: 'PDF für Web optimieren',
    metaDescription: 'PDF für schnelle Web-Anzeige optimieren. Aktivieren Sie das progressive Laden (Fast Web View).',
    keywords: ['pdf linearisieren', 'fast web view', 'pdf optimieren web', 'schnelles laden pdf'],
    description: `
      <p>Linearisieren Sie Ihre Dokumente für eine blitzschnelle Anzeige im Web. Linearisierte PDFs (auch "Fast Web View" genannt) können bereits im Browser angezeigt werden, bevor die gesamte Datei heruntergeladen ist.</p>
      <p>Dies verbessert die Benutzererfahrung bei großen Dokumenten erheblich, da die erste Seite sofort erscheint.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Laden Sie das Dokument hoch, das online veröffentlicht werden soll.' },
      { step: 2, title: 'Optimieren', description: 'Klicken Sie auf "Linearisieren", um die interne Struktur neu zu ordnen.' },
      { step: 3, title: 'Speichern', description: 'Laden Sie das web-optimierte PDF herunter.' },
    ],
    useCases: [
      { title: 'Online-Publikationen', description: 'Optimieren Sie E-Books oder Kataloge für Ihre Website.', icon: 'globe' },
      { title: 'E-Mail-Anhänge', description: 'Sorgen Sie dafür, dass Empfänger Dokumente ohne Verzögerung öffnen können.', icon: 'mail' },
      { title: 'Cloud-Dokumente', description: 'Bessere Performance für Dokumente, die direkt aus der Cloud gelesen werden.', icon: 'cloud' },
    ],
    faq: [
      { question: 'Was ist Linearisierung?', answer: 'Es ist eine spezielle Art der PDF-Strukturierung, die das "Streamen" des Inhalts ermöglicht.' },
      { question: 'Wird die Datei dadurch kleiner?', answer: 'Nicht zwingend, manchmal wird sie minimal größer, lädt aber gefühlt deutlich schneller.' },
      { question: 'Ist das Format kompatibel?', answer: 'Ja, linearisierte PDFs funktionieren in jedem Standard-PDF-Reader.' },
    ],
  },

  'repair-pdf': {
    title: 'PDF reparieren',
    metaDescription: 'Beschädigte PDF-Dateien reparieren. Stellen Sie Inhalte aus defekten Dokumenten wieder her.',
    keywords: ['pdf reparieren', 'defektes pdf retten', 'pdf wiederherstellen', 'corrupted pdf fix'],
    description: `
      <p>Versuchen Sie, beschädigte oder korrupte PDF-Dateien zu retten. Dieses Tool analysiert die Dokumentenstruktur und baut diese neu auf, um so viele Inhalte wie möglich wiederherzustellen.</p>
      <p>Nützlich bei Dateien, die sich nicht öffnen lassen, Fehlermeldungen anzeigen oder beim Download beschädigt wurden.</p>
    `,
    howToUse: [
      { step: 1, title: 'Defektes PDF hochladen', description: 'Wählen Sie die Datei aus, die Fehlermeldungen verursacht.' },
      { step: 2, title: 'Reparatur starten', description: 'Das Tool versucht, die Cross-Reference-Table und die Objektstruktur zu fixen.' },
      { step: 3, title: 'Ergebnis prüfen', description: 'Laden Sie die reparierte Version herunter und prüfen Sie den Inhalt.' },
    ],
    useCases: [
      { title: 'Dateirettung', description: 'Stellen Sie wichtige Daten aus PDFs wieder her, die nicht mehr aufgehen.', icon: 'refresh-cw' },
      { title: 'Fehlerbehebung', description: 'Fixen Sie Dateien, die in manchen Viewern falsch angezeigt werden.', icon: 'wrench' },
      { title: 'Datenwiederherstellung', description: 'Rettungsversuch für unvollständige Downloads oder Übertragungsfehler.', icon: 'file-check' },
    ],
    faq: [
      { question: 'Klappt die Reparatur immer?', answer: 'Der Erfolg hängt vom Grad der Beschädigung ab. Bei schwerem Datenverlust ist eine Rettung leider nicht immer möglich.' },
      { question: 'Bleiben alle Bilder erhalten?', answer: 'Das Tool versucht alles zu retten, bei defekten Objekten kann es jedoch zu Verlusten kommen.' },
      { question: 'Ist mein Original sicher?', answer: 'Ja, Sie arbeiten an einer Kopie; Ihr lokales Original bleibt unverändert.' },
    ],
  },

  // ==================== SECURE PDF ====================
  'encrypt-pdf': {
    title: 'PDF verschlüsseln',
    metaDescription: 'PDF mit Passwort schützen. Fügen Sie Verschlüsselung hinzu und legen Sie Berechtigungen fest.',
    keywords: ['pdf passwort schutz', 'pdf verschlüsseln', 'pdf sichern', 'dokumentenschutz'],
    description: `
      <p>Schützen Sie Ihre vertraulichen Dokumente mit einer starken Verschlüsselung. Legen Sie ein Benutzer-Passwort zum Öffnen und ein Inhaber-Passwort zum Schutz der Berechtigungen fest.</p>
      <p>Wählen Sie zwischen 128-Bit oder 256-Bit AES-Verschlüsselung für höchste Sicherheitsstandards.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie die zu schützende Datei aus.' },
      { step: 2, title: 'Passwörter setzen', description: 'Geben Sie ein starkes Passwort ein und legen Sie fest, was erlaubt ist (z. B. Drucken).' },
      { step: 3, title: 'Sichern', description: 'Laden Sie das verschlüsselte PDF herunter.' },
    ],
    useCases: [
      { title: 'Vertrauliche Daten', description: 'Schützen Sie Finanzberichte oder persönliche Unterlagen.', icon: 'lock' },
      { title: 'Sicherer Versand', description: 'Versenden Sie Verträge geschützt per E-Mail.', icon: 'shield' },
      { title: 'Nutzungskontrolle', description: 'Verhindern Sie, dass Unbefugte Ihre Inhalte kopieren oder drucken.', icon: 'key' },
    ],
    faq: [
      { question: 'Was ist der Unterschied bei den Passwörtern?', answer: 'Das Benutzer-Passwort ist zum Lesen; das Inhaber-Passwort schützt die Rechte-Einstellungen.' },
      { question: 'Wie sicher ist die Verschlüsselung?', answer: 'Wir nutzen AES-Verschlüsselung, die als Industriestandard gilt und extrem sicher ist.' },
      { question: 'Kann ich das Passwort später ändern?', answer: 'Ja, mit dem Inhaber-Passwort können Sie den Schutz jederzeit anpassen oder entfernen.' },
    ],
  },

  'decrypt-pdf': {
    title: 'PDF entsperren',
    metaDescription: 'Passwort aus PDF-Dateien entfernen. Entsperren Sie geschützte Dokumente dauerhaft.',
    keywords: ['pdf passwort entfernen', 'pdf entsperren', 'pdf schutz aufheben', 'pdf decrypt'],
    description: `
      <p>Entfernen Sie den Passwortschutz dauerhaft aus Ihren PDFs. Sobald Sie das korrekte Passwort einmal eingegeben haben, erstellt das Tool eine ungeschützte Kopie für den einfachen Zugriff.</p>
      <p>Hinweis: Sie müssen das aktuelle Passwort kennen. Dieses Tool dient nicht zum Knacken fremder Passwörter.</p>
    `,
    howToUse: [
      { step: 1, title: 'Geschütztes PDF wählen', description: 'Laden Sie die passwortgeschützte Datei hoch.' },
      { step: 2, title: 'Passwort eingeben', description: 'Geben Sie das gültige Passwort ein, um die Datei zu autorisieren.' },
      { step: 3, title: 'Entsperren', description: 'Laden Sie die nun völlig freie PDF-Version herunter.' },
    ],
    useCases: [
      { title: 'Schutz entfernen', description: 'Machen Sie Dokumente für den internen Gebrauch leichter zugänglich.', icon: 'unlock' },
      { title: 'Archivierung', description: 'Entfernen Sie Passwörter vor der Langzeitarchivierung, um Zugriffsprobleme zu vermeiden.', icon: 'archive' },
      { title: 'Workflow-Vereinfachung', description: 'Erstellen Sie offene Kopien für Team-Mitglieder.', icon: 'share-2' },
    ],
    faq: [
      { question: 'Kann das Tool unbekannte Passwörter knacken?', answer: 'Nein, aus Sicherheitsgründen müssen Sie das Passwort besitzen, um es entfernen zu können.' },
      { question: 'Wird das Original geändert?', answer: 'Nein, es wird eine neue, ungeschützte Datei generiert.' },
      { question: 'Gehen Daten verloren?', answer: 'Nein, der Inhalt bleibt exakt gleich, lediglich die Sicherheitsabfrage fällt weg.' },
    ],
  },

  'edit-metadata': {
    title: 'Metadaten bearbeiten',
    metaDescription: 'Dokumenteigenschaften von PDFs ändern. Bearbeiten Sie Titel, Autor, Betreff und Keywords.',
    keywords: ['pdf metadaten ändern', 'pdf autor ändern', 'pdf titel bearbeiten', 'dokumenten-info'],
    description: `
      <p>Passen Sie die versteckten Informationen Ihres PDFs an. Ändern oder löschen Sie Titel, Autor, Thema und Schlagworte, um die Datei professionell zu präsentieren oder Ihre Privatsphäre zu schützen.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Wählen Sie Ihr Dokument aus.' },
      { step: 2, title: 'Felder ausfüllen', description: 'Geben Sie neue Informationen für Titel, Autor und andere Felder ein.' },
      { step: 3, title: 'Speichern', description: 'Laden Sie das PDF mit den aktualisierten Metadaten herunter.' },
    ],
    useCases: [
      { title: 'SEO-Optimierung', description: 'Hinterlegen Sie Keywords und Beschreibungen direkt in der PDF-Datei.', icon: 'search' },
      { title: 'Professionelle Kennzeichnung', description: 'Setzen Sie den korrekten Firmennamen als Autor.', icon: 'user' },
      { title: 'Vorbereitung zur Veröffentlichung', description: 'Sorgen Sie für saubere Dokumenteigenschaften vor dem Upload.', icon: 'file-check' },
    ],
    faq: [
      { question: 'Welche Felder sind änderbar?', answer: 'Titel, Autor, Betreff, Stichwörter, Ersteller und Produzent.' },
      { question: 'Kann ich alle Daten löschen?', answer: 'Ja, Sie können die Felder auch leer lassen, um alle Informationen zu entfernen.' },
      { question: 'Werden auch XMP-Daten geändert?', answer: 'Ja, das Tool aktualisiert sowohl Standard- als auch XMP-Metadaten.' },
    ],
  },

  'pdf-to-pptx': {
    title: 'PDF in PowerPoint',
    metaDescription: 'PDF in PowerPoint (PPTX) Präsentation konvertieren. Jede Seite wird zu einer hochwertigen Folie.',
    keywords: ['pdf in pptx', 'pdf zu powerpoint', 'pdf folien konvertieren', 'pdf präsentation'],
    description: `
      <p>PDF in PowerPoint konvertiert Ihre PDF-Dokumente in bearbeitbare PowerPoint-Präsentationen (PPTX). Jede PDF-Seite wird in eine hochwertige Folie umgewandelt, wobei das visuelle Layout fast perfekt erhalten bleibt.</p>
      <p>Dieses Tool ist ideal für die Umwandlung von Berichten, Handouts oder anderen PDF-Inhalten in ein Präsentationsformat.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Ziehen Sie Ihre PDF-Datei hierher oder klicken Sie zum Auswählen.' },
      { step: 2, title: 'Qualität wählen', description: 'Wählen Sie die Bildqualität (DPI) für die Folien.' },
      { step: 3, title: 'Konvertieren', description: 'Klicken Sie auf Konvertieren, um Ihre PowerPoint-Präsentation zu erstellen.' },
    ],
    useCases: [
      { title: 'Präsentationserstellung', description: 'Wandeln Sie PDF-Berichte in Folien für Meetings um.', icon: 'presentation' },
      { title: 'Schulungsmaterial', description: 'Machen Sie aus PDF-Schulungsunterlagen interaktive PowerPoint-Präsentationen.', icon: 'book-open' },
      { title: 'Inhalte wiederverwenden', description: 'Nutzen Sie bestehende PDF-Inhalte als Basis für neue Präsentationen.', icon: 'refresh-cw' },
    ],
    faq: [
      { question: 'Sind die Folien bearbeitbar?', answer: 'Jede Folie enthält ein Bild der PDF-Seite. Sie können Text und Elemente darüberlegen.' },
      { question: 'Welche DPI sollte ich wählen?', answer: 'Nutzen Sie 150 DPI für Bildschirme und 300 DPI für den Druck.' },
      { question: 'Gehen mehrseitige PDFs?', answer: 'Ja, jede Seite des PDFs wird zu einer separaten Folie.' },
    ],
  },

  'pdf-to-excel': {
    title: 'PDF in Excel',
    metaDescription: 'PDF in Excel-Tabelle konvertieren. Tabellen in XLSX-Format extrahieren.',
    keywords: ['pdf in excel', 'pdf zu xlsx', 'tabellen extrahieren', 'pdf daten'],
    description: `
      <p>PDF in Excel konvertiert Ihre PDF-Dokumente in bearbeitbare Microsoft Excel-Tabellen (XLSX). Das Tool erkennt automatisch Tabellen in Ihrem PDF und extrahiert sie in separate Arbeitsblätter.</p>
      <p>Ideal für die Analyse von Finanzberichten, Rechnungen oder anderen tabellarischen Daten.</p>
    `,
    howToUse: [
      { step: 1, title: 'PDF hochladen', description: 'Laden Sie Ihre PDF-Datei hoch.' },
      { step: 2, title: 'Verarbeiten', description: 'Das Tool identifiziert und extrahiert Tabellen automatisch.' },
      { step: 3, title: 'Excel herunterladen', description: 'Laden Sie die Excel-Datei mit den extrahierten Daten herunter.' },
    ],
    useCases: [
      { title: 'Finanzanalyse', description: 'Konvertieren Sie Kontoauszüge oder Rechnungen für die Analyse.', icon: 'trending-up' },
      { title: 'Datenextraktion', description: 'Holen Sie Daten aus Forschungsberichten in Excel.', icon: 'database' },
      { title: 'Inventarlisten', description: 'Wandeln Sie Bestandslisten von PDF in Tabellen um.', icon: 'clipboard' },
    ],
    faq: [
      { question: 'Wie werden Tabellen behandelt?', answer: 'Erkannte Tabellen werden auf entsprechende Arbeitsblätter in der Excel-Datei verteilt.' },
      { question: 'Was wenn keine Tabellen da sind?', answer: 'Ein Infoblatt wird erstellt, falls keine Tabellen gefunden werden.' },
      { question: 'Bleibt die Formatierung erhalten?', answer: 'Die Daten bleiben erhalten, aber komplexe Formatierungen werden für die Tabelle vereinfacht.' },
    ],
  },
};
