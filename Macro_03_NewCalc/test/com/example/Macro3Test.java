/*
 * Trying to execute tests in a test suite
 *
 */

package com.example;

import com.sun.star.lib.loader.Loader;
import java.lang.reflect.Method;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

/**
 *
 * @author brant
 */
public class Macro3Test {

    public Macro3Test() {
    }

    @Before
    public void setUp() {
    }

    @After
    public void tearDown() {
    }

    @Test
    public void testMacro3() throws Exception {
        String className="com.example.Macro3";
        if ( className != null ) {
            ClassLoader cl = Loader.getCustomLoader();
            Class c = cl.loadClass( className );
            Method m = c.getMethod( "main", new Class[] { String[].class } );
            m.invoke( null, new Object[] { new String[0] } );
        }

    }

}