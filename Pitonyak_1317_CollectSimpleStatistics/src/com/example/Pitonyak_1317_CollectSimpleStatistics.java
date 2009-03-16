/*
 * Pitonyak_1317_CollectSimpleStatistics.java
 *
 * Created on 2009.03.16 - 16:26:22
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XSentenceCursor;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XWordCursor;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Pitonyak_1317_CollectSimpleStatistics {
    
    /** Creates a new instance of Pitonyak_1317_CollectSimpleStatistics */
    public Pitonyak_1317_CollectSimpleStatistics() {
    }
    
    //Sub CollectSimpleStatistics
    //    Dim oCursor
    //    Dim nPars As Long
    //    Dim nSentences As Long
    //    Dim nWords As Long
    //    REM Create a text cursor.
    //    oCursor = ThisComponent.Text.createTextCursor()
    //    oCursor.gotoStart(False)
    //    Do
    //        nPars = nPars + 1
    //    Loop While oCursor.gotoNextParagraph(False)
    //    oCursor.gotoStart(False)
    //    Do
    //        nSentences = nSentences + 1
    //    Loop While oCursor.gotoNextSentence(False)
    //    oCursor.gotoStart(False)
    //    Do
    //        nWords = nWords + 1
    //    Loop While oCursor.gotoNextWord(False)
    //    MsgBox "Paragraphs: " & nPars & CHR$(10) &_
    //    "Sentences: " & nSentences & CHR$(10) &_
    //    "Words: " & nWords & CHR$(10), 0, "Doc Statistics"
    //End Sub

    public static void main(String[] args) {
        Integer nPars = 0;
        Integer nSentences = 0;
        Integer nWords = 0;
        try {
            XComponentContext context = Bootstrap.bootstrap();
            XMultiComponentFactory serviceManager = context.getServiceManager();

            XDesktop desktop= (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context));

            XComponent document = (XComponent) UnoRuntime.queryInterface(XComponent.class,
                desktop.getCurrentComponent());

            XTextDocument textDocument = (XTextDocument) UnoRuntime.queryInterface ( XTextDocument.class,
                document);

            XTextCursor textCursor = textDocument.getText().createTextCursor();
            textCursor.gotoStart(false);
            XParagraphCursor paragraphCursor = (XParagraphCursor)
                UnoRuntime.queryInterface( XParagraphCursor.class, textCursor );
            do {
                nPars++;
            } while (paragraphCursor.gotoNextParagraph(false));
            System.out.println("Paragraphs: "+nPars);

            textCursor.gotoStart(false);
            XSentenceCursor sentenceCursor = (XSentenceCursor)
                UnoRuntime.queryInterface( XSentenceCursor.class, textCursor );
            do {
                nSentences++;
            } while (sentenceCursor.gotoNextSentence(false));
            System.out.println("Sentences: "+nSentences);
            
            textCursor.gotoStart(false);
            XWordCursor wordCursor = (XWordCursor)
                UnoRuntime.queryInterface( XWordCursor.class, textCursor );
            do {
                nWords++;
            } while (wordCursor.gotoNextWord(false));
            System.out.println("Words: "+nWords);

        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}
