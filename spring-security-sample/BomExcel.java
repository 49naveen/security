package com.thomsonreuters.ginger;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.thomsonreuters.suitegtm.specialscheme.billofmaterialapi.SsBillOfMaterialItemApi;
import com.thomsonreuters.suitegtm.specialscheme.billofmaterialapi.SsBillOfMaterialItemFlexApi;
import com.thomsonreuters.suitegtm.specialscheme.billofmaterialapi.SwOrganizationAlternativeItemApi;
import com.thomsonreuters.suitegtm.specialscheme.billofmaterialapi.SwOrganizationApi;
import com.thomsonreuters.suitegtm.specialscheme.billofmaterialapi.SwOrganizationItemApi;
import com.thomsonreuters.suitegtm.specialscheme.billofmaterialapi.SwProductAlternativeItemApi;
import com.thomsonreuters.suitegtm.specialscheme.billofmaterialapi.SwProductApi;
import com.thomsonreuters.suitegtm.specialscheme.billofmaterialapi.SwProductItemApi;

public class BomExcel {
    private static final Logger logger = Logger.getLogger(BomExcel.class);
    private static final String BILL_OF_MATERIAL_SHEET_NAME = "BillOfMaterial";
    private static final String BILL_OF_MATERIAL_ITEM_SHEET_NAME = "BillOfMaterialItem";
    
    private XSSFWorkbook workbook;

    private List<BillOfMaterial> billOfMaterials;
    
    public BomExcel(File file) throws InvalidFormatException, IOException {
        workbook = new XSSFWorkbook(file);
        fetchBillOfMaterialData();
    }
    
    public void close() throws IOException {
        workbook.close();
    }
    
    public List<BillOfMaterial> getBillOfMaterials() {
        return billOfMaterials;
    }

    private void fetchBillOfMaterialData() {
        billOfMaterials = new ArrayList<BillOfMaterial>();
        XSSFSheet sheet = workbook.getSheet(BILL_OF_MATERIAL_SHEET_NAME);
        Map<String, Integer> columnNameToIndexMap = columnNameToIndexMap(sheet);
        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            if (rowNum == 0) {
                continue;
            }
            String recordOperation = fetchCellValue(row, "recordOperation", columnNameToIndexMap);
            String recordStamp = fetchCellValue(row, "recordStamp", columnNameToIndexMap);
            String systemCode = fetchCellValue(row, "systemCode", columnNameToIndexMap);
            String composite = fetchCellValue(row, "composite", columnNameToIndexMap);
            String initialValidity = fetchCellValue(row, "initialValidity", columnNameToIndexMap);
            String finalValidity = fetchCellValue(row, "finalValidity", columnNameToIndexMap);
            String quantity = fetchCellValue(row, "quantity", columnNameToIndexMap);
            String version = fetchCellValue(row, "version", columnNameToIndexMap);
            String measureUnit = fetchCellValue(row, "measureUnit", columnNameToIndexMap);
            String partNumber = fetchCellValue(row, "partNumber", columnNameToIndexMap);
            String organizationCode = fetchCellValue(row, "organizationCode", columnNameToIndexMap);
            String organizationTypeCode = fetchCellValue(row, "organizationTypeCode",
                    columnNameToIndexMap);
            String flexField1 = fetchCellValue(row, "flexField1", columnNameToIndexMap);
            String flexField2 = fetchCellValue(row, "flexField2", columnNameToIndexMap);
            String flexField3 = fetchCellValue(row, "flexField3", columnNameToIndexMap);
            String flexField4 = fetchCellValue(row, "flexField4", columnNameToIndexMap);
            String flexField5 = fetchCellValue(row, "flexField5", columnNameToIndexMap);

            composite += "-" + partNumber.substring(0, 4);

            BillOfMaterial billOfMaterial = new BillOfMaterial();
            billOfMaterial.setRecordOperation(recordOperation);
            billOfMaterial.setRecordStamp(recordStamp);
            billOfMaterial.setSystemCode(systemCode);
            billOfMaterial.setComposite(composite);
            billOfMaterial.setInitialValidity(initialValidity);
            billOfMaterial.setFinalValidity(finalValidity);
            billOfMaterial.setQuantity(quantity);
            billOfMaterial.setVersion(version);
            billOfMaterial.setMeasureUnit(measureUnit);
            SwProductApi product = new SwProductApi();
            product.setPartNumber(partNumber);
            SwOrganizationApi organization = new SwOrganizationApi();
            organization.setOrganizationCode(organizationCode);
            organization.setOrganizationTypeCode(organizationTypeCode);
            product.setOrganization(organization);
            billOfMaterial.setBillOfMaterialProduct(product);
            billOfMaterial.setFlexField1(flexField1);
            billOfMaterial.setFlexField2(flexField2);
            billOfMaterial.setFlexField3(flexField3);
            billOfMaterial.setFlexField4(flexField4);
            billOfMaterial.setFlexField5(flexField5);
            setDefaultValuesIfBlank(billOfMaterial);

            fetchBillOfMaterialItemData(billOfMaterial);
            billOfMaterials.add(billOfMaterial);
        }
    }
    
    private void fetchBillOfMaterialItemData(BillOfMaterial billOfMaterial) {
        List<SsBillOfMaterialItemApi> items = new ArrayList<SsBillOfMaterialItemApi>();
        XSSFSheet sheet = workbook.getSheet(BILL_OF_MATERIAL_ITEM_SHEET_NAME);
        Map<String, Integer> columnNameToIndexMap = columnNameToIndexMap(sheet);
        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            if (rowNum == 0) {
                continue;
            }
            String model = fetchCellValue(row, "model", columnNameToIndexMap);
            if (!StringUtils.equals(billOfMaterial.getBillOfMaterialProduct().getPartNumber(),
                    model)) {
                continue;
            }
            String recordOperation = fetchCellValue(row, "recordOperation", columnNameToIndexMap);
            String recordStamp = fetchCellValue(row, "recordStamp", columnNameToIndexMap);
            String partNumber = fetchCellValue(row, "partNumber", columnNameToIndexMap);
            String partNumberAlternative = fetchCellValue(row, "partNumberAlternative",
                    columnNameToIndexMap);
            String organizationCode = fetchCellValue(row, "organizationCode", columnNameToIndexMap);
            String organizationTypeCode = fetchCellValue(row, "organizationTypeCode",
                    columnNameToIndexMap);

            String item = fetchCellValue(row, "item", columnNameToIndexMap);
            String itemQuantity = fetchCellValue(row, "itemQuantity", columnNameToIndexMap);
            String required = fetchCellValue(row, "required", columnNameToIndexMap);
            String lossPercentage = fetchCellValue(row, "lossPercentage", columnNameToIndexMap);
            String minQtyMaterial = fetchCellValue(row, "minQtyMaterial", columnNameToIndexMap);
            String maxQtyMaterial = fetchCellValue(row, "maxQtyMaterial", columnNameToIndexMap);
            String externalSystemDate = fetchCellValue(row, "externalSystemDate",
                    columnNameToIndexMap);
            String externalSystemIdentifier = fetchCellValue(row, "externalSystemIdentifier",
                    columnNameToIndexMap);
            String bomMaterialType = fetchCellValue(row, "bomMaterialType", columnNameToIndexMap);
            String initialValidityItem = fetchCellValue(row, "initialValidityItem",
                    columnNameToIndexMap);
            String finalValidityItem = fetchCellValue(row, "finalValidityItem",
                    columnNameToIndexMap);
            String measureUnit = fetchCellValue(row, "measureUnit", columnNameToIndexMap);
            String alternativeMeasureUnit = fetchCellValue(row, "alternativeMeasureUnit",
                    columnNameToIndexMap);
            String alternativeQuantity = fetchCellValue(row, "alternativeQuantity",
                    columnNameToIndexMap);

            String itemFlexField1 = fetchCellValue(row, "itemFlexField1", columnNameToIndexMap);
            String itemFlexField2 = fetchCellValue(row, "itemFlexField2", columnNameToIndexMap);
            String itemFlexField3 = fetchCellValue(row, "itemFlexField3", columnNameToIndexMap);
            String itemFlexField4 = fetchCellValue(row, "itemFlexField4", columnNameToIndexMap);
            String itemFlexField5 = fetchCellValue(row, "itemFlexField5", columnNameToIndexMap);

            SsBillOfMaterialItemApi itemApi = new SsBillOfMaterialItemApi();
            itemApi.setRecordOperation(recordOperation);
            itemApi.setRecordStamp(recordStamp);
            itemApi.setBillOfMaterialItemProduct(new SwProductItemApi());
            itemApi.getBillOfMaterialItemProduct().setPartNumber(partNumber);
            SwOrganizationItemApi org = new SwOrganizationItemApi();
            org.setOrganizationCode(organizationCode);
            org.setOrganizationTypeCode(organizationTypeCode);
            itemApi.getBillOfMaterialItemProduct().setOrganization(org);
            itemApi.setBillOfMaterialItemProductAlternative(new SwProductAlternativeItemApi());
            itemApi.getBillOfMaterialItemProductAlternative().setPartNumber(partNumberAlternative);
            SwOrganizationAlternativeItemApi orgAlt = new SwOrganizationAlternativeItemApi();
            orgAlt.setOrganizationCode(organizationCode);
            orgAlt.setOrganizationTypeCode(organizationTypeCode);
            itemApi.getBillOfMaterialItemProductAlternative().setOrganization(orgAlt);

            itemApi.setItem(item);
            itemApi.setItemQuantity(itemQuantity);
            itemApi.setRequired(required);
            if (StringUtils.isBlank(lossPercentage)) {
                lossPercentage = "1";
            }
            itemApi.setLossPercentage(lossPercentage);
            itemApi.setMinQtyMaterial(minQtyMaterial);
            itemApi.setMaxQtyMaterial(maxQtyMaterial);
            itemApi.setExternalSystemDate(externalSystemDate);
            itemApi.setExternalSystemIdentifier(externalSystemIdentifier);
            itemApi.setBomMaterialType(bomMaterialType);
            itemApi.setInitialValidityItem(initialValidityItem);
            itemApi.setFinalValidityItem(finalValidityItem);
            itemApi.setMeasureUnit(measureUnit);
            itemApi.setAlternativeMeasureUnit(alternativeMeasureUnit);
            itemApi.setAlternativeQuantity(alternativeQuantity);

            itemApi.setFlexField1(itemFlexField1);
            itemApi.setFlexField2(itemFlexField2);
            itemApi.setFlexField3(itemFlexField3);
            itemApi.setFlexField4(itemFlexField4);
            itemApi.setFlexField5(itemFlexField5);

            SsBillOfMaterialItemFlexApi itemFlex = new SsBillOfMaterialItemFlexApi();
            for (int i = 1; i <= 43; i++) {
                String productFlexFieldName = "flexField" + i;
                String productFlexFieldValue = fetchCellValue(row, productFlexFieldName,
                        columnNameToIndexMap);
                String methodName = "set" + Character.toUpperCase(productFlexFieldName.charAt(0))
                        + productFlexFieldName.substring(1);
                try {
                    Class<?> c = Class.forName(
                            "com.thomsonreuters.suitegtm.specialscheme.billofmaterialapi.SsBillOfMaterialItemFlexApi");
                    Method method = c.getMethod(methodName, String.class);
                    method.invoke(itemFlex, productFlexFieldValue);
                } catch (NoSuchMethodException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                } catch (SecurityException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                } catch (ClassNotFoundException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                } catch (IllegalArgumentException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
            for (int i = 1; i <= 4; i++) {
                String productFlexFieldName = "flexFieldLong" + i;
                String productFlexFieldValue = fetchCellValue(row, productFlexFieldName,
                        columnNameToIndexMap);
                String methodName = "set" + Character.toUpperCase(productFlexFieldName.charAt(0))
                        + productFlexFieldName.substring(1);
                try {
                    Class<?> c = Class.forName(
                            "com.thomsonreuters.suitegtm.specialscheme.billofmaterialapi.SsBillOfMaterialItemFlexApi");
                    Method method = c.getMethod(methodName, String.class);
                    method.invoke(itemFlex, productFlexFieldValue);
                } catch (NoSuchMethodException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                } catch (SecurityException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                } catch (ClassNotFoundException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                } catch (IllegalArgumentException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
            itemApi.setBillOfMaterialFlexField(itemFlex);

            setDefaultValuesIfBlank(itemApi);
            items.add(itemApi);
        }
        billOfMaterial.getSsBillOfMaterialItemApi().addAll(items);
    }
    
    private String fetchCellValue(Row row, String columnName,
            Map<String, Integer> columnNameToIndexMap) {
        Integer columnIndex = columnNameToIndexMap.get(columnName);
        /*
         * if(logger.isDebugEnabled()) { logger.debug("fetchCellValue: " +
         * columnName + " - " + columnIndex); }
         */
        Cell cell = row.getCell(columnIndex);
        String value = "";
        if (cell != null) {
            cell.setCellType(Cell.CELL_TYPE_STRING);
            value = cell.getStringCellValue().trim();
        }
        return value;
    }

    private Map<String, Integer> columnNameToIndexMap(XSSFSheet sheet) {
        Map<String, Integer> columnNameToIndexMap = new HashMap<String, Integer>();
        Row firstRow = sheet.getRow(sheet.getFirstRowNum());
        for (Cell cell : firstRow) {
            String cellValue = cell.getStringCellValue();
            int columnIndex = cell.getColumnIndex();
            if(logger.isDebugEnabled()) {
                logger.debug(cellValue + " - " + columnIndex);
            }
            columnNameToIndexMap.put(cellValue, columnIndex);
        }
        return columnNameToIndexMap;
    }

    private void setDefaultValuesIfBlank(BillOfMaterial billOfMaterial) {
        if (StringUtils.isBlank(billOfMaterial.getRecordOperation())) {
            billOfMaterial.setRecordOperation("CREATE-UPDATE");
        }
        if (StringUtils.isBlank(billOfMaterial.getRecordStamp())) {
            billOfMaterial.setRecordStamp("2015-11-05T09:00:00.00001");
        }
        if (StringUtils.isBlank(billOfMaterial.getBillOfMaterialProduct().getOrganization()
                .getOrganizationCode())) {
            billOfMaterial.getBillOfMaterialProduct().getOrganization().setOrganizationCode("561");
        }
        if (StringUtils.isBlank(billOfMaterial.getBillOfMaterialProduct().getOrganization()
                .getOrganizationTypeCode())) {
            billOfMaterial.getBillOfMaterialProduct().getOrganization()
                    .setOrganizationTypeCode("OGT002");
        }
    }

    private void setDefaultValuesIfBlank(SsBillOfMaterialItemApi billOfMaterialItem) {
        if (StringUtils.isBlank(billOfMaterialItem.getRecordOperation())) {
            billOfMaterialItem.setRecordOperation("CREATE-UPDATE");
        }
        if (StringUtils.isBlank(billOfMaterialItem.getRecordStamp())) {
            billOfMaterialItem.setRecordStamp("2015-11-05T09:00:00.00001");
        }
        if (StringUtils.isBlank(billOfMaterialItem.getBillOfMaterialItemProduct().getOrganization()
                .getOrganizationCode())) {
            billOfMaterialItem.getBillOfMaterialItemProduct().getOrganization()
                    .setOrganizationCode("561");
        }
        if (StringUtils.isBlank(billOfMaterialItem.getBillOfMaterialItemProduct().getOrganization()
                .getOrganizationTypeCode())) {
            billOfMaterialItem.getBillOfMaterialItemProduct().getOrganization()
                    .setOrganizationTypeCode("OGT002");
        }
        if (StringUtils.isBlank(billOfMaterialItem.getBillOfMaterialItemProductAlternative()
                .getOrganization().getOrganizationCode())) {
            billOfMaterialItem.getBillOfMaterialItemProductAlternative().getOrganization()
                    .setOrganizationCode("561");
        }
        if (StringUtils.isBlank(billOfMaterialItem.getBillOfMaterialItemProductAlternative()
                .getOrganization().getOrganizationTypeCode())) {
            billOfMaterialItem.getBillOfMaterialItemProductAlternative().getOrganization()
                    .setOrganizationTypeCode("OGT002");
        }
        if (StringUtils.isBlank(billOfMaterialItem.getExternalSystemIdentifier())) {
            billOfMaterialItem.setExternalSystemIdentifier("HIOTS");
        }
    }
    
}
