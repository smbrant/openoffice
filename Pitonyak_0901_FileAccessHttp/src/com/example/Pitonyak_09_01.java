/*
 * Pitonyak_09_01.java
 *
 * Created on 2009.02.04 - 08:07:32
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.ucb.XSimpleFileAccess;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Pitonyak_09_01 {
    
    /** Creates a new instance of Pitonyak_09_01 */
    public Pitonyak_09_01() {
    }
    
    //ch09 lis01
    //Sub ManagerCreatesAService
    //    Dim vFileAccess
    //    Dim s As String
    //    Dim vManager
    //    vManager = GetProcessServiceManager()
    //    vFileAccess = vManager.CreateInstance("com.sun.star.ucb.SimpleFileAccess")
    //    s = vFileAccess.getContentType("http://www.pitonyak.org/AndrewMacro.sxw")
    //    Print s
    //End Sub
    public static void main(String[] args) {
        try {
            XComponentContext context = Bootstrap.bootstrap();

            XMultiComponentFactory srvMan = context.getServiceManager();

            XSimpleFileAccess vFileAccess =
                    (XSimpleFileAccess) UnoRuntime.queryInterface( XSimpleFileAccess.class,
                    srvMan.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", context));

            String s= vFileAccess.getContentType("http://www.pitonyak.org/AndrewMacro.sxw");
            System.out.println("ContentType: "+s);
            System.exit( 0 );
        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}
