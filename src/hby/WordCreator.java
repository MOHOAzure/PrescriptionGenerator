package hby;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;

import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import medic.Prescription;

public class WordCreator {

	public WordCreator(){
		super();
	}
	
    public void createPrescriptionWord(ArrayList<Prescription> prescriptionList, String departeDate){
        try (XWPFDocument doc = new XWPFDocument()) {
        	
            XWPFParagraph p1 = doc.createParagraph();
            p1.setAlignment(ParagraphAlignment.LEFT);
            
            //�۩w�q���
            CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
            if (sectPr == null) sectPr = doc.getDocument().getBody().addNewSectPr();
            CTPageMar pageMar = sectPr.getPgMar();
            if (pageMar == null) pageMar = sectPr.addNewPgMar();
            
            int margin=567; // ~1cm
            //e.g., 720 TWentieths of an Inch Point (Twips) = 720/20 = 36 pt = 36/72 = 0.5"            
            pageMar.setLeft(BigInteger.valueOf(margin));
            pageMar.setRight(BigInteger.valueOf(margin));
            pageMar.setTop(BigInteger.valueOf(2*margin));
            pageMar.setBottom(BigInteger.valueOf(2*margin));
            //pageMar.setFooter(BigInteger.valueOf(720));
            //pageMar.setHeader(BigInteger.valueOf(720));
            //pageMar.setGutter(BigInteger.valueOf(0));
            
            int p_count = 0;

            for(Prescription p : prescriptionList){
            	p_count++;// the number of prescription
            	/*
            	XWPFRun head_space = p1.createRun();
            	head_space.setFontSize(28);
            	head_space.addBreak();
            	*/
	            XWPFRun r1 = p1.createRun();
	            r1.setText("�x���a���`��| �C�ʯf�s��B���� ���ĳq����");
	            r1.setFontSize(26);
	            r1.setBold(true);
	            r1.setFontFamily("�з���");
	            r1.addBreak();
	            
	            XWPFRun r2 = p1.createRun();
	            r2.setText(departeDate);
	            r2.setFontSize(26);
	            r2.setBold(true);
	            r2.setFontFamily("�з���");
	            r2.addBreak();

	            XWPFRun r3 = p1.createRun();
	            r3.setText("���Ыe�@��U�� 3 �I�e�N���O IC �d��ܫO���ե~��ǡ�");
	            r3.setFontSize(21);
	            r3.setBold(true);
	            r3.setFontFamily("�з���");
	            r3.addBreak();

	            XWPFRun r4 = p1.createRun();
	            r4.setFontSize(28);
	            r4.setBold(true);
	            r4.setFontFamily("�з���");
	            r4.setText("��O�G\t"+p.getHouseName());
	            r4.addBreak();
	            r4.setText("�m�W�G\t"+p.getResidentName());
	            r4.addBreak();
	            r4.setText("��O�G\t"+p.getDepartmentName());
	            r4.addBreak();
	            r4.setText("���Ħ��O�G\t�� "+p.getProgress()+" ��");
	            r4.addBreak();

	            XWPFRun r5 = p1.createRun();
	            r5.setFontSize(24);
	            r5.setBold(true);
	            r5.setFontFamily("�з���");
	            r5.setText("���Щ��ѤU�� 1 �I��ܫO���ը���  ����!");
	            
	            //�@�����W�U�ⳡ��
	            if(p_count%2!=0){
	            	XWPFRun a_little_break = p1.createRun();
	            	a_little_break.setFontSize(32);
	            	a_little_break.addBreak();
	            	a_little_break.addBreak();
	            	a_little_break.addBreak();
	            	a_little_break.addBreak();
	            }
	            
	            // �C��ӺC�ҳq����N����
	            if(p_count%2==0 && p_count<prescriptionList.size())
	            	r5.addBreak(BreakType.PAGE);
            }
            
	        //create word file
	        try (FileOutputStream out = new FileOutputStream("�C�ҳq����.docx")) {
	            doc.write(out);
	            out.close();
	            doc.close();
	        }//End of XWPFDocument doc
        } catch (IOException e) {
			e.printStackTrace();
		}
        //System.out.println("Prescription Word file is created");
    }
}
