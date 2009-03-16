/*
 * Pitonyak_09_03.java
 *
 * Created on 2009.02.02 - 16:17:30
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.lang.XMultiComponentFactory;

/**
 *
 * @author brant
 */
public class Pitonyak_09_03 {
    
    /** Creates a new instance of Pitonyak_09_03 */
    public Pitonyak_09_03() {
    }

    //ch09 lis03
    //Sub HowManyServicesSupported
    //    Dim vManager
    //    Dim sServices
    //    vManager = GetProcessServiceManager()
    //    sServices = vManager.getAvailableServiceNames()
    //    Print "Service manager supports ";UBound(sServices);" services"
    //End Sub
    public static void main(String[] args) throws Exception {

        XComponentContext context = Bootstrap.bootstrap();
        XMultiComponentFactory srvMan = context.getServiceManager();

        String[] sServices= srvMan.getAvailableServiceNames();
        for (String service : sServices){
            System.out.println(service);
        }
        System.out.println("Total of services: "+sServices.length);
        System.exit( 0 );
    }
    
}
