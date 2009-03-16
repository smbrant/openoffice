/*
 * Macro_14_ClearBlankLines2.java
 *
 * Created on 2009.02.16 - 17:17:22
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;

/**
 *
 * @author brant
 */
public class Macro_14_ClearBlankLines2 {
    
    /** Creates a new instance of Macro_14_ClearBlankLines2 */
    public Macro_14_ClearBlankLines2() {
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            // get the remote office component context
            XComponentContext context = Bootstrap.bootstrap();
            Writer writer= new Writer(context);
            writer.clearBlankLines();
        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}
