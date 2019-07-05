package com.github.feyond.excel;

import lombok.Getter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * @author
 * @create 2017-09-30 12:31
 **/
public class ImportExcel {

    /**
     * 工作薄对象
     */
    @Getter
    private Workbook wb;

    public ImportExcel(InputStream inputStream) throws IOException, InvalidFormatException {
        this.wb = WorkbookFactory.create(inputStream);
    }

    public ImportExcel(File file) throws IOException, InvalidFormatException {
        if (!file.exists()) {
            throw new RuntimeException("导入文档为空!");
        } else if (!file.isFile() || !file.canRead()) {
            throw new RuntimeException("导入文档不可读!");
        } else {
            this.wb = WorkbookFactory.create(new FileInputStream(file));
        }
    }

    public ImportExcel(String filepath) throws IOException, InvalidFormatException {
        this(new File(filepath));
    }

    public <E> List<E> getDataList(int sheetIndex, int dataRow, Class<E> cls, int... groups) {
        AnnotationSheetWrapper sheetWrapper = new AnnotationSheetWrapper(this, sheetIndex);
        return sheetWrapper.getDataList(dataRow, cls, groups);
    }

    public List<Map<Integer, Object>> getDataList(int sheetIndex, int dataRow) {
        SheetWrapper sheet = new SheetWrapper(this, sheetIndex);
        return sheet.getDataList(dataRow);
    }

    public <E> List<E> getDataListWithHeader(int sheetIndex, int headerRow, Class<E> cls, int... groups) {
        AnnotationSheetWrapper sheetWrapper = new AnnotationSheetWrapper(this, sheetIndex);
        return sheetWrapper.getDataListWithHeader(headerRow, cls, groups);
    }

    public List<Map<String, Object>> getDataListWithHeader(int sheetIndex, int headerRow) {
        SheetWrapper sheet = new SheetWrapper(this, sheetIndex);
        return sheet.getDataListWithHeader(headerRow);
    }
}
