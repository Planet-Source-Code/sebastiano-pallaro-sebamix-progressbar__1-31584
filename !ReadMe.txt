SebaMix ProgressBar

v. 1.1.3

~~~~~~~~~
Copyright
~~~~~~~~~

This DLL is (c) by Pallaro Sebastiano. 
You can use it as you wish (but first send me an e-mail! :) ).
The DLL is give "as is", so the author don't take any kind of responsability.

~~~~~~~~~~~~
ENUMERATIONS
~~~~~~~~~~~~

 Public Enum SMOrientation
     SMHorizontal = 0
     SMVertical = 1
 End Enum

 Public Enum SMDrawStyles
     SM_MergePen = vbMergePen
     SM_MergePenNot = vbMergePenNot
     SM_MergeNotPen = vbMergeNotPen
     SM_NotXorPen = vbNotXorPen
     SM_MaskPen = vbMaskPen
     SM_NotMaskPen = vbNotMaskPen
     SM_XorPen = vbXorPen
     SM_Invert = vbInvert
     SM_MaskPenNot = vbMaskPenNot
     SM_MaskNotPen = vbMaskNotPen
     SM_NotMergePen = vbNotMergePen
 End Enum

~~~~~~~~~
PROPERIES
~~~~~~~~~

 · BackColor
    (OLE_COLOR) The back color of the ProgressBar;

 · Caption
    (String) Set the progressbar's caption. In this property
     is left blank, the progress value will be show;

 · ForeColor
    (OLE_COLOR) The color of the caption;

 · Max
    (Long) The max value that the Value property can assume;

 · Min
    (Long) The min value that the value property can assume;

 · Percent
    (Integer) Read only, return the curren value's percent;

 · PictureProgress
    (SMDrawStyles) Specify the print style of the progression.
     Using this propery will change the DrawMode of the PictureBox. Refer
     to the OldDrawMode propery to have the original DrawMode.

 · ProgressColor
    (OLE_COLOR) The color of the progression;

 · OldDrawMode
    (Byte) Read only, return the original DrawMode propery of the
     PictureBox;

 · Orientation
    (SMOrientation) Set the orientation for the progressbar
    (horizontal or vertical)

 · ShowCaption
    (Boolean) If true the caption will be show;

 · TextAfter
    (String) The string that will print before the value propery on the 
     caption if the caption propery is empty;
     (e.g. "Step 40/100")

 · TextBefore
    (String) The string that will be print after the value property on the
     caption if the caption property is empty;
     (e.g. "40/100 done")

 · TextMiddle
    (string) Stringa che verrà stampata tra la Value raggiunta e Max
    nella caption se la proprietà Caption è vuota e UsePercent è Falsa;
    (e.g. the "/" of "200/300");

 · UsePercent
    (Boolean) If true, the progress percent will be show;

 · Value
    (Long) The value of the progression;
    IMPORTANT : the progressbar's picture will be print only when the
    value property is set;

~~~~~~~
METHODS
~~~~~~~
 · InitPB (ByVal myPictureBox As Object, Optional myOrientation As SMOrientation)
    Initialize the progressbar object. You must give the picturebox object and the
    orientation (default horizontal).

 · GetVersion () As String
    Return the version (Major.Minor.Revision);

 · GiveOfficeBorder()
    Give the office border to the progressbar;

 · AboutBox()
    Open the About-Box;

~~~~~~
EVENTS
~~~~~~
 · Progress (Value As Long, Percent As Integer)
    Raised when the "value" property is set (when the picture redraw is terminated).
    It pass the value and the percent of the progressbar;

~~~~
Note (SORRY, ITALIAN ONLY)
~~~~
 · Il metodo InitPB setta le seguenti proprietà della PictureBox
    passatagli:
    - Picture.AutoRedraw = True;
    - Picture.ScaleMode = vbTwips;
    Si consiglia di non modificare in alcun modo le properietà
    della PictureBox per non ottenere effetti indesiderati;
 · Valori di default:
    - Inizializzazione della classe:
        Caption = ""
        UsePercent = False
        ShowCaption = True
        Min = 0
        Max = 100
        Value = 0
        Percent = 0
    - Metodo InitPB:
        BackColor = RGB(150, 150, 150)
        ProgressColor = RGB(100, 100, 250)
        ForeColor = vbWhite
        Orientation = myOrientation (se non passato vale SMHorizontal - 0)

~~~~~~~~
Utilizzo (SORRY, ITALIAN ONLY)
~~~~~~~~

 Includere SMPB tra i riferimenti del progetto (menù progetto>riferimenti).
 Se non si è registrata la DLL includerla utilizzando il pulsante "Sfoglia"
 Se necessario registrare la DLL attraverso il comando REGSVR32.
 Ove la progressbar è necessaria includere questa dichiarazione (preferibilmente
 creare la variabile globale a livello di maschera - o modulo);

 Private WithEvents SMPB1 As smProgressBar

 La clausola WithEvents implica che gli eventi siano abilitati.
 Per inizializzare la maschera inserire le seguenti chiamate (preferibilmente
 nel Form_Load)

 Set SMPB1 = New smProgressBar
 SMPB1.InitPB PictureBox, SMVertical

 Quando la ProgressBar non serve più (per esempio nel Form_Load) è
 consigliabile inserire la seguente chiamata:

 Set SMPB1 = Nothing

 Può capitare infatti che il programma possa andare in GPF (Global Protection
 Fail) se si lasciano degli oggetti istanziati in memoria.

~~~~~~~~~~~~~~~~~~~~~
Storia delle Versioni (SORRY, ITALIAN ONLY)
~~~~~~~~~~~~~~~~~~~~~
 · 1.0.0 : Versione di partenza. 
           Bachi conosciuti: errata rappresentazione della progressione 
            nel caso la proprietà Min sia maggiore di 0;
           [Versione non distribuita]
 · 1.0.1 : Classe iniziale portata in una DLL (SMPB.DLL);
           [Versione non distribuita]
 · 1.0.2 : Risolto il problema con la proprietà Min;
           Bachi conosciuti : la percentuale impazzisce quando la proprietà
            Min è maggiore di 0;
           [Versione non distribuita]
 · 1.0.3 : Aggiunto lo scroll verticale;
           Bachi conosciuti : lo scorrimento verticale va dall'alto verso il
            basso e non viceversa; la caption resta orizzontale;
           [Versione non distribuita]
 · 1.0.4 : Corretto l'orientamento dello scroll verticale, aggiunto il
            metodo GetVersion e la proprietà ProgressColor; reimpostati i
            colori di default ed eliminato lo sfarfallio riscontrato
            durante la progressione;
           [Versione non distribuita]
 · 1.0.5 : Aggiunto l'evento Progress; Gestito l'evento Class_Terminate per
            il corretto scarico dalla memoria degli oggetti istanziati;
           [Versione non distribuita]
 · 1.1.0 : Prima versione distribuita.
           Bachi conosciuti : la caption resta orrizzontale anche durante lo
            scorrimento verticale;
           [Versione distribuita]
 · 1.1.1 : Aggiunto il metodo GiveOfficeBorder, che però ha un effetto 
            irreversibile; aggiunto il metodo AboutBox; risolto il problema 
            che sorgeva quando Value corrispondeva a Min (veniva visualizzata 
            comunque una piccola progressione nella PictureBox);
           [Versione non distribuita]
 · 1.1.2 : Aggiunta la possibilità di visualizzare un'immagine di sfondo nella
            ProgressBar (con la possibilità di applicare vari effetti alla
            barra di progressione invece del semplice colore impostato; Aggiunta
            la proprietà OldDrawMode per un restore del DrawMode iniziale della
            PictureBox; risolto il problema delle proprietà BackColor, ForeColor, 
            ProgressColor che non si potevano settare per un errore interno;
           [Versione non distribuita]
 · 1.1.3 : Aggiunte le proprietà TextBefore, TextAfter e TextMiddle;
           [Versione distribuita]
