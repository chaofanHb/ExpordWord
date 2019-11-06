package hncr.bzPlatform.prjMgr.word;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.net.MalformedURLException;
import java.nio.charset.Charset;
import java.util.HashMap;
import java.util.Map;

import freemarker.core.ParseException;
import freemarker.template.Configuration;
import freemarker.template.MalformedTemplateNameException;
import freemarker.template.Template;
import freemarker.template.TemplateNotFoundException;

/**
 * @Description:word导出帮助类 通过freemarker模板引擎来实现
 * @author:LiaoFei
 * @date :2016-3-24 下午3:49:25
 * @version V1.0
 * 
 */
public class WordGeneratorWithFreemarker {
	private static Configuration configuration = null;

	static {
		configuration = new Configuration(Configuration.VERSION_2_3_23);
		configuration.setDefaultEncoding("utf-8");
		configuration.setClassicCompatible(true);
		configuration.setClassForTemplateLoading(
				WordGeneratorWithFreemarker.class,
				"/hncr/bzPlatform/prjMgr/word/ftl");

	}

	private WordGeneratorWithFreemarker() {

	}

	public static void createDoc(Map<String, Object> dataMap,String templateName, OutputStream out)throws Exception {
		Template t = configuration.getTemplate(templateName);
		t.setEncoding("utf-8");
		WordHtmlGeneratorHelper.handleAllObject(dataMap);

		try {
			Writer w = new OutputStreamWriter(out,Charset.forName("utf-8"));
			t.process(dataMap, w);
			w.close();
		} catch (Exception ex) {
			ex.printStackTrace();
			throw new RuntimeException(ex);
		}
	}

	public static void main(String[] args) throws Exception {
		HashMap<String, Object> data = new HashMap<String, Object>();

		StringBuilder sb = new StringBuilder();
		sb.append("<div>");
		sb.append("<img style='height:100px;width:200px;display:block;' src='G:\\aaa.png' />");
		sb.append("<img style='height:100px;width:200px;display:block;' src='G:\\29-1Z414195R3.jpg' />");
		sb.append("</br><span>中国梦，幸福梦！</span>");
		sb.append("</div>");
		RichHtmlHandler handler = new RichHtmlHandler(sb.toString());

		handler.setDocSrcLocationPrex("file:///C:/268BA2D4");
		handler.setDocSrcParent("test.files");
		handler.setNextPartId("01D593ED.81F97840");
		handler.setShapeidPrex("_x56fe__x7247__x0020");
		handler.setSpidPrex("_x0000_i");
		handler.setTypeid("#_x0000_t75");

		handler.handledHtml(false);

		String bodyBlock = handler.getHandledDocBodyBlock();
		System.out.println("bodyBlock:\n"+bodyBlock);

		String handledBase64Block = "";
		if (handler.getDocBase64BlockResults() != null
				&& handler.getDocBase64BlockResults().size() > 0) {
			for (String item : handler.getDocBase64BlockResults()) {
				handledBase64Block += item + "\n";
			}
		}
		data.put("imagesBase64String", handledBase64Block);

		String xmlimaHref = "";
		if (handler.getXmlImgRefs() != null
				&& handler.getXmlImgRefs().size() > 0) {
			for (String item : handler.getXmlImgRefs()) {
				xmlimaHref += item + "\n";
			}
		}
		
		StringBuilder sb1 = new StringBuilder();
		sb1.append("<img style='height:100px;width:200px;display:block;' src='G:\\3_boneix.jpg' />1111111");
		RichHtmlHandler handler1 = new RichHtmlHandler(sb1.toString());
		handler1.setDocSrcLocationPrex("file:///C:/268BA2D4");
		handler1.setDocSrcParent("test.files");
		handler1.setNextPartId("01D593ED.81F97840");
		handler1.setShapeidPrex("_x56fe__x7247__x0020");
		handler1.setSpidPrex("_x0000_i");
		handler1.setTypeid("#_x0000_t75");
		handler1.handledHtml(false);
		String bodyBlock1 = handler1.getHandledDocBodyBlock();
		String handledBase64Block1 = "";
		if (handler1.getDocBase64BlockResults() != null
				&& handler1.getDocBase64BlockResults().size() > 0) {
			for (String item : handler1.getDocBase64BlockResults()) {
				handledBase64Block1 += item + "\n";
			}
		}
		String xmlimaHref1 = "";
		if (handler1.getXmlImgRefs() != null
				&& handler1.getXmlImgRefs().size() > 0) {
			for (String item : handler1.getXmlImgRefs()) {
				xmlimaHref1 += item + "\n";
			}
		}
		data.put("imagesBase64String1", handledBase64Block1);
		data.put("imagesXmlHrefString1", xmlimaHref1);
		
		data.put("imagesXmlHrefString", xmlimaHref);
		data.put("name", "张三");
		data.put("content1", bodyBlock);
		data.put("content2", bodyBlock1);

		String docFilePath = "d:\\temp.doc";
		System.out.println(docFilePath);
		File f = new File(docFilePath);
		OutputStream out;
		try {
			out = new FileOutputStream(f);
			WordGeneratorWithFreemarker.createDoc(data, "test.mht", out);

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (MalformedTemplateNameException e) {
			e.printStackTrace();
		} catch (ParseException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
}