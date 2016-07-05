/*
 * Info: Name=Lo,WeiShun 
 * Author: raliclo
 * Filename: myVB2JS.java
 * Date and Time: Jul 5, 2016 4:53:50 AM
 * Project Name: myvb2js
 */
 /*
 * Copyright 2016 raliclo.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package myVB2JS;

import com.google.vb2js.VbaJsConverter;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;

/**
 *
 * @author raliclo
 */
public class myVB2JS {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // setup variable
        args = new String[1];
        args[0] = "./test/test.vb";
        ArrayList<String> vbCode = new ArrayList<>();

        System.out.println("[VB] Before Conversion");
        if (args.length == 1) {
            try (BufferedReader br = new BufferedReader(new FileReader(args[0]))) {

                String sCurrentLine;

                while ((sCurrentLine = br.readLine()) != null) {
                    System.out.println(sCurrentLine);
                    vbCode.add(sCurrentLine);
                }

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        System.out.println("\n---[VB -> JS] Conversion---\n");
        System.out.println("[JS] After Conversion");
        System.out.println(VbaJsConverter.convert(vbCode));
    }
}
