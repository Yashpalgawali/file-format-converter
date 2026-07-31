package com.example.demo.service;

import java.io.InputStream;

public interface IExcelToPdfService {
	
	public byte[] convertExcelToPdf(InputStream excelStream) throws Exception;
	
}
