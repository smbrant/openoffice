/*
 * Pitonyak_1328_RemoveEmptySpace.java
 *
 * Created on 2009.02.13 - 10:43:54
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;

/**
 *
 * @author brant
 */
public class Pitonyak_1328_RemoveEmptySpace {
    
    /** Creates a new instance of Pitonyak_1328_RemoveEmptySpace */
    public Pitonyak_1328_RemoveEmptySpace() {
    }

    //Sub RemoveEmptySpace
    //    Dim oCursors(), i%
    //    If Not CreateSelectedTextIterator(ThisComponent, _
    //        "ALL empty space will be removed from the ENTIRE document?", oCursors()) Then Exit Sub
    //    For i% = LBOUND(oCursors()) To UBOUND(oCursors())
    //        RemoveEmptySpaceWorker (oCursors(i%, 0), oCursors(i%, 1), ThisComponent.Text)
    //    Next i%
    //End Sub
    public static void main(String[] args) {
        try {
            XComponentContext context = Bootstrap.bootstrap();
        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}
