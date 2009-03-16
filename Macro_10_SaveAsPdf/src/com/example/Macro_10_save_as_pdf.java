/*
 * Macro_10_save_as_pdf.java
 *
 * Created on 2009.02.06 - 15:35:16
 *
 * Save the current document in pdf format
 *
 */

package com.example;

import com.sun.star.beans.PropertyValue;
import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Macro_10_save_as_pdf {
    
    /** Creates a new instance of Macro_10_save_as_pdf */
    public Macro_10_save_as_pdf() {
    }

// from http://codesnippets.services.openoffice.org/Writer/Writer.StoreActualDocumentAs__pdf.snip

///* StoreActualAsPDF.rex */
///* Macro */
///* How to store an opened file as *.pdf */
//
///* get the script context */
//xScriptContext=uno.getScriptContext()
//
///* get the document and the component context */
//oDoc=xScriptContext~getDocument
//oContext=xScriptContext~getComponentContext
//
///* defining the storing properties */
//xStorable = oDoc~XStorable
//
//storeprops = bsf.createArray(.UNO~propertyValue, 2)
//storeprops[1] = .UNO~PropertyValue~new
//storeprops[1]~Name  = "FilterName"
//storeprops[1]~Value = "writer_pdf_Export"
//storeprops[2] = .UNO~PropertyValue~new
//storeprops[2]~Name  = "CompressMode"
//storeprops[2]~Value = 2
//
///* retrieve url from document */
//url=oDoc~getURL()
//
///* if document has been saved to an url */
//if (url<>"") then do
//	parse var url url "." .
//	url= url || ".pdf"
//	xStorable~storeToUrl(url, storeprops)
//end
//
///* if document has not been saved */
//else do
//	.bsf.dialog~messagebox("File has to be saved first.")
//end
//
//::requires UNO.CLS
//
    public static void main(String[] args) {
        try {
            XComponentContext context = Bootstrap.bootstrap();
            XMultiComponentFactory serviceManager = context.getServiceManager();
            XDesktop desktop =
              (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context));
            XTextDocument document= (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, desktop.getCurrentComponent());
            XStorable storable= (XStorable) UnoRuntime.queryInterface(XStorable.class, document);

            PropertyValue storeprops[] = new PropertyValue[2];
            storeprops[0]= new PropertyValue();
            storeprops[0].Name= "FilterName";
            storeprops[0].Value= "writer_pdf_Export";
            storeprops[1]= new PropertyValue();
            storeprops[1].Name= "CompressMode";
            storeprops[1].Value= 2;

            String url= document.getURL();
            if (!url.equals("")){
                url= url.substring(0, url.lastIndexOf("."))+".pdf";
                storable.storeToURL(url, storeprops);
            }else{
                System.out.println("File has to be saved first.");
            }
        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}
