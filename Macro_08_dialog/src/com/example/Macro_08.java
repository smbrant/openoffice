/*
 * Macro_08.java
 *
 * Created on 2009.02.04 - 16:10:49
 *
 * A dialog built in java
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.awt.PushButtonType;
import com.sun.star.awt.XTextComponent;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Macro_08 {
    
    /** Creates a new instance of Macro_08 */
    public Macro_08() {
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            // get the remote office component context
            XComponentContext context = Bootstrap.bootstrap();
            Dialog dialog = new Dialog(context,200,200,80,280,"Send Mail","sendMailDialog");

            dialog.insertFixedText(10, 10, 30, "Subject:");
            XTextComponent subjectField = dialog.insertTextField(42, 10, 228, "Subject...");
            dialog.insertFixedText(10, 30, 30, "Recipients:");
            XTextComponent recipientsField = dialog.insertTextField(42, 30, 228, "Recipients...");

            dialog.insertButton(105, 60, 30, "OK", PushButtonType.OK_value);
            dialog.insertButton(145, 60, 30, "Cancel", PushButtonType.CANCEL_value);

            short returnValue = dialog.execute();
            System.out.println("Subject: "+subjectField.getText());
            System.out.println("Recipients: "+recipientsField.getText());
            System.out.println("Return: "+returnValue);

            XMultiComponentFactory xMCF = context.getServiceManager();
            XDesktop xDesktop =
              (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  xMCF.createInstanceWithContext("com.sun.star.frame.Desktop",context));

            XTextDocument xtextdoc = (XTextDocument)
                UnoRuntime.queryInterface(XTextDocument.class, xDesktop.getCurrentComponent());
            if (xtextdoc!=null){
                if (returnValue==1){
                    XText xtext = xtextdoc.getText();
                    XTextCursor xtextcursor = xtext.createTextCursor();
                    xtext.insertString(xtextcursor, "Subject: "+subjectField.getText()+"\n"+
                                                    "Recipients: "+recipientsField.getText()+"\n\n",
                            false);
                }
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
