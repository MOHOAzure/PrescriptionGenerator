/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package hby;

public class Main {

    public static void main(String[] args){
    	String excelFileName = args[0];
    	String date = args[1];
    	
    	createPrescriptionWord(excelFileName, date);
    	
    	//getSheetNames(excelFileName);
    }
    
    private static void getSheetNames(String excelFileName) {
		ExcelReader er = new ExcelReader();
		er.getSheetNames(excelFileName);		
	}

	public static void createPrescriptionWord(String excelFileName, String date){
		ExcelReader er = new ExcelReader();
		WordCreator wc = new WordCreator();
		
		wc.createPrescriptionWord(
							er.readPrescriptionExcel(excelFileName, date), 
							er.getDepartDate());		
    }
}