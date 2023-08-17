package org.example;



import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
    public static void main(String[] args) {
        String excelFilePath = "/Users/seekekrishna/Downloads/Mahendras1.xlsx"; // Replace with the actual path to your Excel file

        FileInputStream inputStream = null;

        try {
            inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream);

            // Replace "Android" with the sheet name where your data resides
            Sheet sheet = workbook.getSheet("Android");

            // Column indexes where data needs to be populated
            int itemIDColumnIndex = 1;
            int itemNameColumnIndex = 2;
            int itemBrandColumnIndex = 3;
            int priceColumnIndex = 4;
            int indexColumnIndex = 5;
            int centerColumnIndex = 6;
            int timeSlotColumnIndex = 7;
            int batchStartDateColumnIndex = 8;
            int validityColumnIndex = 9;
            int modeOfLearningColumnIndex = 10;
            int itemListNameColumnIndex = 11;
            int eventActionColumnIndex = 12; // Assuming it's the 12th column (0-based index)
            int eventLabelColumnIndex = 13;
            int orgIdColumnIndex = 14;
            int accountIdColumnIndex = 15;
            int userIdColumnIndex = 16;
            int primaryGoalColumnIndex = 17;
            int primaryExamColumnIndex = 18;
            int examCategoryColumnIndex = 19;
            int deviceInfoColumnIndex = 20;
            int platformColumnIndex = 21;
            int localeColumnIndex = 22;
            int moduleNameColumnIndex = 23;
            int screenNameColumnIndex = 24;
            int featureNameColumnIndex = 25;
            int clientIdColumnIndex = 26;
            int contentIdColumnIndex = 27;
            int contentTypeColumnIndex = 28;
            int contentSubtypeColumnIndex = 29;
            int titleColumnIndex = 30;
            int vleCodeColumnIndex = 31;
            int errorTypeColumnIndex = 32;
            int errorCodeColumnIndex = 33;
            int pajIdColumnIndex = 34;
            int pajSessionIdColumnIndex = 35;
            int certificationTestStateColumnIndex = 36;
            int pajTotalStepsColumnIndex = 37;
            int videoDownloadedColumnIndex = 38;
            int searchTermColumnIndex = 39;
            int searchSourceColumnIndex = 40;
            int searchPathColumnIndex = 41;
            int verticalRankColumnIndex = 42;
            int filterTypeColumnIndex = 43;
            int platformSessionIdColumnIndex = 44;
            int schoolIdColumnIndex = 45;
            int slotIdColumnIndex = 46;
            int classIdColumnIndex = 47;
            int teacherIdColumnIndex = 48;
            int videoSourceColumnIndex = 49;
            int watchTimeColumnIndex = 50;
            int videoLengthColumnIndex = 51;
            int nameColumnIndex = 52;
            int transactionIdColumnIndex = 53;
            int taxColumnIndex = 54;
            int shippingColumnIndex = 55;
            int booksAndBagsChargesColumnIndex = 56;
            int currencyColumnIndex = 57;
            int valueColumnIndex = 58;
            int couponColumnIndex = 59;








            for (Row row : sheet) {
                // Assuming "Logs" is in the first column (0-based index)
                Cell logsCell = row.getCell(0);

                if (logsCell != null && logsCell.getCellType() == Cell.CELL_TYPE_STRING) {
                    String logData = logsCell.getStringCellValue();

                    // Parse the log data to extract individual details
                    String itemID = extractValueByKey(logData, "item_id");
                    String itemName = extractValueByKey(logData, "item_name");
                    String itemBrand = extractValueByKey(logData, "item_brand");
                    String price = extractValueByKey(logData, "price");
                    String index = extractValueByKey(logData, "index");
                    String center = extractValueByKey(logData, "center");
                    String timeSlot = extractValueByKey(logData, "time_slot");
                    String batchStartDate = extractValueByKey(logData, "batch_start_date");
                    String validity = extractValueByKey(logData, "validity");
                    String modeOfLearning = extractValueByKey(logData, "mode_of_learning");
                    String itemListName = extractValueByKey(logData, "item_list_name");
                    String eventAction = extractValueByKey(logData, "eventAction");
                    String eventLabel = extractValueByKey(logData, "eventLabel");
                    String orgId = extractValueByKey(logData, "org_id");
                    String accountId = extractValueByKey(logData, "account_id");
                    String userId = extractValueByKey(logData, "userId");
                    String primaryGoal = extractValueByKey(logData, "primary_goal");
                    String primaryExam = extractValueByKey(logData, "primary_exam");
                    String examCategory = extractValueByKey(logData, "exam_category");
                    String deviceInfo = extractValueByKey(logData, "device_info");
                    String platform = extractValueByKey(logData, "platform");
                    String locale = extractValueByKey(logData, "locale");
                    String moduleName = extractValueByKey(logData, "module_name");
                    String screenName = extractValueByKey(logData, "screen_name");
                    String featureName = extractValueByKey(logData, "feature_name");
                    String clientId = extractValueByKey(logData, "client_id");
                    String contentId = extractValueByKey(logData, "content_id");
                    String contentType = extractValueByKey(logData, "content_type");
                    String contentSubtype = extractValueByKey(logData, "content_subtype");
                    String title = extractValueByKey(logData, "title");
                    String vleCode = extractValueByKey(logData, "vle_code");
                    String errorType = extractValueByKey(logData, "error_type");
                    String errorCode = extractValueByKey(logData, "error_code");
                    String pajId = extractValueByKey(logData, "paj_id");
                    String pajSessionId = extractValueByKey(logData, "paj_session_id");
                    String certificationTestState = extractValueByKey(logData, "certification_test_state");
                    String pajTotalSteps = extractValueByKey(logData, "paj_total_steps");
                    String videoDownloaded = extractValueByKey(logData, "video_downloaded");
                    String searchTerm = extractValueByKey(logData, "search_term");
                    String searchSource = extractValueByKey(logData, "search_source");
                    String searchPath = extractValueByKey(logData, "search_path");
                    String verticalRank = extractValueByKey(logData, "vertical_rank");
                    String filterType = extractValueByKey(logData, "filter_type");
                    String platformSessionId = extractValueByKey(logData, "platform_session_id");
                    String schoolId = extractValueByKey(logData, "school_id");
                    String slotId = extractValueByKey(logData, "slot_id");
                    String classId = extractValueByKey(logData, "class_id");
                    String teacherId = extractValueByKey(logData, "teacher_id");
                    String videoSource = extractValueByKey(logData, "video_source");
                    String watchTime = extractValueByKey(logData, "Watch_Time");
                    String videoLength = extractValueByKey(logData, "Video_Length");
                    String name = extractValueByKey(logData, "name");
                    String transactionId = extractValueByKey(logData, "transaction_id");
                    String tax = extractValueByKey(logData, "tax");
                    String shipping = extractValueByKey(logData, "shipping");
                    String booksAndBagsCharges = extractValueByKey(logData, "books_and_bags_charges");
                    String currency = extractValueByKey(logData, "currency");
                    String value = extractValueByKey(logData, "value");
                    String coupon = extractValueByKey(logData, "coupon");









                    // Populate the data in the respective columns
                    row.createCell(itemIDColumnIndex).setCellValue(itemID);
                    row.createCell(itemNameColumnIndex).setCellValue(itemName);
                    row.createCell(itemBrandColumnIndex).setCellValue(itemBrand);
                    row.createCell(priceColumnIndex).setCellValue(price);
                    row.createCell(indexColumnIndex).setCellValue(index);
                    row.createCell(centerColumnIndex).setCellValue(center);
                    row.createCell(timeSlotColumnIndex).setCellValue(timeSlot);
                    row.createCell(batchStartDateColumnIndex).setCellValue(batchStartDate);
                    row.createCell(validityColumnIndex).setCellValue(validity);
                    row.createCell(modeOfLearningColumnIndex).setCellValue(modeOfLearning);
                    row.createCell(itemListNameColumnIndex).setCellValue(itemListName);

                    row.createCell(eventActionColumnIndex).setCellValue(eventAction);
                    row.createCell(eventLabelColumnIndex).setCellValue(eventLabel);
                    row.createCell(orgIdColumnIndex).setCellValue(orgId);
                    row.createCell(accountIdColumnIndex).setCellValue(accountId);
                    row.createCell(userIdColumnIndex).setCellValue(userId);
                    row.createCell(primaryGoalColumnIndex).setCellValue(primaryGoal);
                    row.createCell(primaryExamColumnIndex).setCellValue(primaryExam);
                    row.createCell(examCategoryColumnIndex).setCellValue(examCategory);
                    row.createCell(deviceInfoColumnIndex).setCellValue(deviceInfo);
                    row.createCell(platformColumnIndex).setCellValue(platform);
                    row.createCell(localeColumnIndex).setCellValue(locale);
                    row.createCell(moduleNameColumnIndex).setCellValue(moduleName);
                    row.createCell(screenNameColumnIndex).setCellValue(screenName);
                    row.createCell(featureNameColumnIndex).setCellValue(featureName);
                    row.createCell(clientIdColumnIndex).setCellValue(clientId);
                    row.createCell(contentIdColumnIndex).setCellValue(contentId);
                    row.createCell(contentTypeColumnIndex).setCellValue(contentType);
                    row.createCell(contentSubtypeColumnIndex).setCellValue(contentSubtype);
                    row.createCell(titleColumnIndex).setCellValue(title);
                    row.createCell(vleCodeColumnIndex).setCellValue(vleCode);
                    row.createCell(errorTypeColumnIndex).setCellValue(errorType);
                    row.createCell(errorCodeColumnIndex).setCellValue(errorCode);
                    row.createCell(pajIdColumnIndex).setCellValue(pajId);
                    row.createCell(pajSessionIdColumnIndex).setCellValue(pajSessionId);
                    row.createCell(certificationTestStateColumnIndex).setCellValue(certificationTestState);
                    row.createCell(pajTotalStepsColumnIndex).setCellValue(pajTotalSteps);
                    row.createCell(videoDownloadedColumnIndex).setCellValue(videoDownloaded);
                    row.createCell(searchTermColumnIndex).setCellValue(searchTerm);
                    row.createCell(searchSourceColumnIndex).setCellValue(searchSource);
                    row.createCell(searchPathColumnIndex).setCellValue(searchPath);
                    row.createCell(verticalRankColumnIndex).setCellValue(verticalRank);
                    row.createCell(filterTypeColumnIndex).setCellValue(filterType);
                    row.createCell(platformSessionIdColumnIndex).setCellValue(platformSessionId);
                    row.createCell(schoolIdColumnIndex).setCellValue(schoolId);
                    row.createCell(slotIdColumnIndex).setCellValue(slotId);
                    row.createCell(classIdColumnIndex).setCellValue(classId);
                    row.createCell(teacherIdColumnIndex).setCellValue(teacherId);
                    row.createCell(videoSourceColumnIndex).setCellValue(videoSource);
                    row.createCell(watchTimeColumnIndex).setCellValue(watchTime);
                    row.createCell(videoLengthColumnIndex).setCellValue(videoLength);
                    row.createCell(nameColumnIndex).setCellValue(name);
                    row.createCell(transactionIdColumnIndex).setCellValue(transactionId);
                    row.createCell(taxColumnIndex).setCellValue(tax);
                    row.createCell(shippingColumnIndex).setCellValue(shipping);
                    row.createCell(booksAndBagsChargesColumnIndex).setCellValue(booksAndBagsCharges);
                    row.createCell(currencyColumnIndex).setCellValue(currency);
                    row.createCell(valueColumnIndex).setCellValue(value);
                    row.createCell(couponColumnIndex).setCellValue(coupon);



                }
            }

            // Save the changes back to the Excel file
            try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
                workbook.write(outputStream);
            }

            System.out.println("Data populated successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                // Close the FileInputStream to release the resources properly
                if (inputStream != null) {
                    inputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    // Function to extract values from the log data using regex
    private static String extractValueByKey(String logData, String key) {
        String pattern = key + "=([^,]+)";
        Pattern regexPattern = Pattern.compile(pattern);
        Matcher matcher = regexPattern.matcher(logData);
        if (matcher.find()) {
            return matcher.group(1);
        }
        return null;
    }
}
