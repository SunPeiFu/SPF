package com.viewhigh.oes.hr.hrbase.bo.poi.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Properties;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import com.viewhigh.oes.common.commondc.utils.PathUtil;

public class MergeDocUtil {

	public static String mergeDocx(List<String> filePathList) throws Exception {

		if (filePathList == null || filePathList.isEmpty()) {
			System.out.println("MergeDocUtil data is null");
			return null;
		}

		InputStream in1 = null;
		InputStream in2 = null;
		OPCPackage src1Package = null;
		OPCPackage src2Package = null;
		String mergePath = getRootPath() + File.separator
				+ "上海公卫合并输出文档out.docx";
		// 最终输出文档
		OutputStream dest = new FileOutputStream(mergePath);
		// 起始遍历索引,指定list的第一个元素
		in1 = new FileInputStream(filePathList.get(0));
		src1Package = OPCPackage.open(in1);
		XWPFDocument src1Document = new XWPFDocument(src1Package);
		// 设置分隔符
		src1Document.createParagraph().setPageBreak(true);
		CTBody src1Body = src1Document.getDocument().getBody();
		try {
			// 第一个已经渲染完成,从第二个开始
			for (int i = 1; i < filePathList.size(); i++) {
				in2 = new FileInputStream(filePathList.get(i));
				src2Package = OPCPackage.open(in2);
				XWPFDocument src2Document = new XWPFDocument(src2Package);
				src2Document.createParagraph().setPageBreak(true);
				CTBody src2Body = src2Document.getDocument().getBody();
				appendBody(src1Body, src2Body);
				in2.close();
				src2Package.close();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}

		src1Document.write(dest);
		src1Package.close();
		return mergePath;
	}

	private static void appendBody(CTBody src, CTBody append) throws Exception {
		XmlOptions optionsOuter = new XmlOptions();
		optionsOuter.setSaveOuter();
		String appendString = append.xmlText(optionsOuter);
		String srcString = src.xmlText();
		String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
		String mainPart = srcString.substring(srcString.indexOf(">") + 1,
				srcString.lastIndexOf("<"));
		String sufix = srcString.substring(srcString.lastIndexOf("<"));
		String addPart = appendString.substring(appendString.indexOf(">") + 1,
				appendString.lastIndexOf("<"));
		CTBody makeBody = CTBody.Factory.parse(prefix + mainPart + addPart
				+ sufix);
		src.set(makeBody);
	}

	// 系统文件相对路径
	public static String getRootPath() {
		return PathUtil.getUtil().getWebInfPath()
				+ "conf/hr/hrbase/poitemplate";
	}

	public static String getPoiValue(String key) throws Exception {
		String value= "";
		Properties prop = new Properties();
		String path= PathUtil.getUtil().getWebInfPath()
		+ "conf/hr/hrbase/poitemplate/workname.properties";
		try {
            prop.load(new FileInputStream(path));
			value = (String) prop.get(key);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return value;
	}
}
