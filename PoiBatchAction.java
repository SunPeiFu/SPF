package com.viewhigh.oes.hr.hrbase.bo.poi;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Map.Entry;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.struts2.interceptor.ServletRequestAware;
import org.apache.struts2.interceptor.ServletResponseAware;
import org.springframework.beans.BeansException;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;

import com.alibaba.fastjson.JSON;
import com.aspose.words.Document;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import com.opensymphony.xwork2.ActionSupport;
import com.viewhigh.oes.hr.hrbase.bo.poi.config.Configure;
import com.viewhigh.oes.hr.hrbase.bo.poi.util.MergeDocUtil;
import com.viewhigh.oes.hr.hrbase.bo.poi.util.PoiUtil;
import com.viewhigh.oes.hr.hrbase.bo.poi.vo.PoiVo;
import com.viewhigh.oes.hr.hrbase.bo.print.FormPrintBO;
import com.viewhigh.oes.hr.hrbase.dao.print.FormPrintDAO;

public class PoiBatchAction extends ActionSupport implements
		ServletRequestAware, ServletResponseAware, ApplicationContextAware {
	private ApplicationContext context;
	private HttpServletRequest request;
	private HttpServletResponse response;
	// 处理渲染Pdf
	private static InputStream license;
	private static InputStream fileInput;
	private static File outputFile;
	// 注入Bo
	private FormPrintBO printBo;
	
	public FormPrintBO getPrintBo() {
		return printBo;
	}

	public void setPrintBo(FormPrintBO printBo) {
		this.printBo = printBo;
	}

	// 合并打印预览
	@SuppressWarnings("all")
	public void previewMergePoi2Pdf() throws Exception {
		
		String mergeDocx = "";
		try {
			request.setCharacterEncoding("UTF-8");
			String params = request.getParameter("params");
			Map<String, Object> para = (Map<String, Object>) JSON.parse(params);
			String empId = para.get("empId").toString();
			// 最大数据集
			Map<String, Object> maxData = printBo.queryPrintAll(empId);
			Map<String, Object> data = maxData;
			// 创建业务模型
			PoiVo vo = new PoiVo();
			String type = "docToPdf";
			// 最终生成的docxList
			List<String> docxList = new ArrayList<String>();
			Map<String, Object> linkedMap = sortMap(data);
			data = linkedMap;
			// 结果返回集
			Map<String, Object> result = linkedMap;
			// 顺序遍历,迭代删除
			Iterator<Entry<String, Object>> it = data.entrySet().iterator();
			while (it.hasNext()) {
				Entry<String, Object> entry = it.next();
				String docxOutPath = chooseTemplate(data, result, vo, type);
				docxList.add(docxOutPath);
				it.remove();
			}
			// 合并模板,生成最终文档
			mergeDocx = MergeDocUtil.mergeDocx(docxList);
			docx2Pdf(mergeDocx);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			File mergeDocxFile = new File(mergeDocx);
			if (mergeDocxFile.exists() && mergeDocxFile.isFile()) {
				mergeDocxFile.delete();
			}
		}
	}
	
	// 合并打印下载
	@SuppressWarnings("all")
	public void printMergePoi2Pdf() throws Exception {
		request.setCharacterEncoding("UTF-8");
		String params = request.getParameter("params");
		Map<String, Object> para = (Map<String, Object>) JSON.parse(params);
		String empId = para.get("empId").toString();
		Map<String, Object> maxData = printBo.queryPrintAll(empId);
		Map<String, Object> data = maxData;
		String mergeDocx ="";
		List<String> docxList = new ArrayList<String>();
		try{
			// 创建业务模型
			PoiVo vo = new PoiVo();
			// 最终生成的docxList
			Map<String, Object> linkedMap = sortMap(data);
			data = linkedMap;
			// 结果返回集
			Map<String, Object> result = linkedMap;
			// 顺序遍历,迭代删除
			Iterator<Entry<String, Object>> it = data.entrySet().iterator();
			while (it.hasNext()) {
				Entry<String, Object> entry = it.next();
				String docxOutPath = chooseTemplate(data, result, vo, null);
				docxList.add(docxOutPath);
				it.remove();
			}
			// 合并模板,生成最终文档
			mergeDocx = MergeDocUtil.mergeDocx(docxList);
			docxList.add(mergeDocx);
			downLoadFile(new File(mergeDocx), vo);
		}catch (Exception e) {
			e.printStackTrace();
		} finally {
			// 删除所有生成临时文件
			for (String filePath : docxList) {
				File file = new File(filePath);
				if (file.exists() && file.isFile()) {
					file.delete();
				}
			}
			
		}
	}

	private Map<String, Object> sortMap(Map<String, Object> data) {
		// 结果有序,用linkedMap
		Map<String, Object> linkedMap = new LinkedHashMap<String, Object>();
		linkedMap.put("empinfo", data.get("empinfo"));
		linkedMap.put("posinfo", data.get("posinfo"));
		linkedMap.put("operatecert", data.get("operatecert"));
		linkedMap.put("postain", data.get("postain"));
		linkedMap.put("yeartain", data.get("yeartain"));
		linkedMap.put("jobedu", data.get("jobedu"));
		linkedMap.put("printcopy", data.get("printcopy"));
		return linkedMap;
	}

	// 选择模板
	private String chooseTemplate(Map<String, Object> data,
			Map<String, Object> result, PoiVo vo, String type) throws Exception {
		Configure config = null;
		// 模板名称
		String templateName = "";
		// 执行具体实现
		result = PoiUtil.executeShangHai(data, result, vo);
		templateName = (String) result.get("templateName");
		config = (Configure) result.get("config");

		String docxOutPath = outFile(result, vo, type, config, templateName);
		return docxOutPath;
	}

	private void docx2Pdf(String docxOutPath) throws IOException {
		File outPdfFile = null;
		BufferedInputStream bis = null;
		BufferedOutputStream bos = null;
		FileOutputStream fileOS = null;
		try {
			// docx2pdf,读取license.xml,此xml为专用license
			String licensePath = MergeDocUtil.getRootPath() + File.separator
					+ "poi-license.xml";
			String pdfPath = docxOutPath.replace(".docx", ".pdf");
			if (!getLicense(licensePath, docxOutPath, pdfPath)) {
				System.out.println("license出错");
				return;
			}
			//  加上注释,测试提交代码是否好使
			Document doc = new Document(fileInput);
			fileOS = new FileOutputStream(outputFile);
			doc.save(fileOS, SaveFormat.PDF);
			fileOS.flush();
			fileOS.close();
			// 获取pdf文件
			outPdfFile = new File(pdfPath);
			if (outPdfFile.exists() && outPdfFile.isFile()
					&& outPdfFile.length() > 0) {
				System.out.println("文件存在");
				// 文件长度>0
				response.setContentType("application/pdf;charset=UTF-8");
				bis = new BufferedInputStream(new FileInputStream(outPdfFile));
				bos = new BufferedOutputStream(response.getOutputStream());
				byte[] buff = new byte[2048];
				int bytesRead;
				while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
					bos.write(buff, 0, bytesRead);
				}
				bos.flush();
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {

			if (bis != null) {
				bis.close();
			}
			if (bos != null) {
				bos.close();
			}
			if (license != null) {
				license.close();
			}
			if (fileInput != null) {
				fileInput.close();
			}

			if (outPdfFile.exists() && outPdfFile.isFile()
					&& outPdfFile.length() > 0) {
				outPdfFile.delete();
			}
		}

	}

	private String outFile(Map<String, Object> result, PoiVo vo, String type,
			Configure config, String templateName)
			throws FileNotFoundException, IOException {
		vo.setM(result);
		String templatePath = MergeDocUtil.getRootPath() + File.separator
				+ templateName + ".docx";
		XWPFTemplate template = XWPFTemplate.compile(templatePath, config)
				.render(vo);
		String docxOutPath = templatePath.replace(".docx", "out.docx");
		File docxOutFile = new File(docxOutPath);
		FileOutputStream out = new FileOutputStream(docxOutFile);
		template.write(out);
		out.flush();
		if ("downLoad".equals(type)) {
			downLoadFile(docxOutFile, vo);
		}
		if (out != null) {
			out.close();
		}
		return docxOutPath;
	}

	// 下载文件
	private void downLoadFile(File docxOutFile, PoiVo vo) throws IOException {
		BufferedInputStream bis = null;
		BufferedOutputStream bos = null;
		try {

			String filePath = docxOutFile.getPath();
			String fileName = filePath
					.substring((filePath.lastIndexOf("\\") + 1));
			// 拼接下载文件默认名
			String empName = "";
			String empCode = "";
			Map<String, Object> basicInfo = vo.getM();
			if (!basicInfo.isEmpty()) {
				empName = (String) (basicInfo.get("empName") == null ? ""
						: basicInfo.get("empName"));
				empCode = (String) (basicInfo.get("empCode") == null ? ""
						: basicInfo.get("empCode"));
			}
			fileName = fileName.replace("out.docx", new StringBuilder().append(
					"_").append(empName).append("_").append(empCode).append(
					".docx"));
			String agent = request.getHeader("User-Agent").toUpperCase();
			if ((agent.indexOf("MSIE") > 0)
					|| ((agent.indexOf("RV") != -1) && (agent
							.indexOf("FIREFOX") == -1))) {
				fileName = URLEncoder.encode(fileName, "UTF-8");
			} else {
				fileName = new String(fileName.getBytes("UTF-8"), "ISO8859-1");
			}
			response.setContentType("application/x-msdownload;");
			response.setHeader("Content-disposition", "attachment; filename="
					+ fileName);
			response.setHeader("Content-Length", String.valueOf(docxOutFile
					.length()));
			bis = new BufferedInputStream(new FileInputStream(docxOutFile));
			bos = new BufferedOutputStream(response.getOutputStream());
			byte[] buff = new byte[2048];
			int bytesRead=0;
			while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
				bos.write(buff, 0, bytesRead);
			}
			bos.flush();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (bis != null) {
				bis.close();
			}
			if (bos != null) {
				bos.close();
			}
			if (docxOutFile.exists() && docxOutFile.isFile()) {
				boolean result = docxOutFile.delete();
				if (result) {
					System.out.println("删除文件成功");
				}
			}
		}
	}

	// 获取License
	public static boolean getLicense(String licensePath, String sourcePath,
			String outPath) {
		boolean result = false;
		try {
			license = new FileInputStream(new File(licensePath));
			// 输出文件
			fileInput = new FileInputStream(new File(sourcePath));
			// 待处理的文件
			outputFile = new File(outPath);

			License aposeLic = new License();
			aposeLic.setLicense(license);
			result = true;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}

	public ApplicationContext getContext() {
		return context;
	}

	public HttpServletRequest getRequest() {
		return request;
	}

	public HttpServletResponse getResponse() {
		return response;
	}

	@Override
	public void setServletRequest(HttpServletRequest request) {
		this.request = request;
	}

	@Override
	public void setServletResponse(HttpServletResponse response) {
		this.response = response;

	}

	@Override
	public void setApplicationContext(ApplicationContext context)
			throws BeansException {
		this.context = context;
	}

	public static InputStream getLicense() {
		return license;
	}

	public static void setLicense(InputStream license) {
		PoiBatchAction.license = license;
	}

	public static InputStream getFileInput() {
		return fileInput;
	}

	public static void setFileInput(InputStream fileInput) {
		PoiBatchAction.fileInput = fileInput;
	}

	public static File getOutputFile() {
		return outputFile;
	}

	public static void setOutputFile(File outputFile) {
		PoiBatchAction.outputFile = outputFile;
	}

	public void setContext(ApplicationContext context) {
		this.context = context;
	}

	public void setRequest(HttpServletRequest request) {
		this.request = request;
	}

	public void setResponse(HttpServletResponse response) {
		this.response = response;
	}

}
