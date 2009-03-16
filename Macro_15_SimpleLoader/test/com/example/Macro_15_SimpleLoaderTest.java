/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

import com.sun.star.lib.loader.Loader;
import org.junit.After;
import org.junit.Before;

/**
 *
 * @author brant
 */
public class Macro_15_SimpleLoaderTest {

    public Macro_15_SimpleLoaderTest() {
    }

    @Before
    public void setUp() throws Exception {
        Loader.load("com.example.Macro_15_SimpleLoader");
    }

    @After
    public void tearDown() {
    }

}