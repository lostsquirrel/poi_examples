/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */
package demo.poi.image;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Demonstrates how to add pictures in a .docx document
 */
public class SimpleImages {

	private static final Logger log = LoggerFactory.getLogger(SimpleImages.class);

	public static void main(String[] args) throws IOException, InvalidFormatException {
		// writeImageToDoc(args);
		InputStream file = SimpleImages.class.getClassLoader().getResourceAsStream("images.docx");
		extractImagesFromWord(file);
	}

	static void writeImageToDoc(String[] args)
			throws InvalidFormatException, IOException, FileNotFoundException {
		XWPFDocument doc = new XWPFDocument();
		XWPFParagraph p = doc.createParagraph();

		XWPFRun r = p.createRun();

		for (String imgFile : args) {
			int format;

			if (imgFile.endsWith(".emf"))
				format = XWPFDocument.PICTURE_TYPE_EMF;
			else if (imgFile.endsWith(".wmf"))
				format = XWPFDocument.PICTURE_TYPE_WMF;
			else if (imgFile.endsWith(".pict"))
				format = XWPFDocument.PICTURE_TYPE_PICT;
			else if (imgFile.endsWith(".jpeg") || imgFile.endsWith(".jpg"))
				format = XWPFDocument.PICTURE_TYPE_JPEG;
			else if (imgFile.endsWith(".png"))
				format = XWPFDocument.PICTURE_TYPE_PNG;
			else if (imgFile.endsWith(".dib"))
				format = XWPFDocument.PICTURE_TYPE_DIB;
			else if (imgFile.endsWith(".gif"))
				format = XWPFDocument.PICTURE_TYPE_GIF;
			else if (imgFile.endsWith(".tiff"))
				format = XWPFDocument.PICTURE_TYPE_TIFF;
			else if (imgFile.endsWith(".eps"))
				format = XWPFDocument.PICTURE_TYPE_EPS;
			else if (imgFile.endsWith(".bmp"))
				format = XWPFDocument.PICTURE_TYPE_BMP;
			else if (imgFile.endsWith(".wpg"))
				format = XWPFDocument.PICTURE_TYPE_WPG;
			else {
				System.err.println("Unsupported picture: " + imgFile
						+ ". Expected emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg");
				continue;
			}

			r.setText(imgFile);
			r.addBreak();
			r.addPicture(new FileInputStream(imgFile), format, imgFile, Units.toEMU(200), Units.toEMU(200)); // 200x200
																												// pixels
			r.addBreak(BreakType.PAGE);
		}

		FileOutputStream out = new FileOutputStream("images.docx");
		doc.write(out);
		out.close();
		// doc.close();
	}

	public static List<byte[]> extractImagesFromWord(InputStream file) {
		try {
			List<byte[]> result = new ArrayList<byte[]>();
			XWPFDocument doc = new XWPFDocument(file);
			log.debug("{}", doc);
			
			for (XWPFPictureData picture : doc.getAllPictures()) {
				result.add(picture.getData());
				log.debug("{}", picture.getFileName());
			}

			return result;
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

}