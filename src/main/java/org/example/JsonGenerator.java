//package org.example;
//
//
//
//    import javax.swing.*;
//import javax.swing.border.TitledBorder;
//import java.awt.*;
//import java.awt.event.ActionEvent;
//import java.awt.event.ActionListener;
//import java.io.FileWriter;
//import java.io.IOException;
//import java.util.UUID;
//import java.util.Random;
//
//    import org.json.JSONException;
//    import org.json.JSONObject;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import java.io.FileOutputStream;
//
//    public class JsonGenerator extends JFrame implements ActionListener {
//
//        private JComboBox<String> moduleDropdown;
//        private JComboBox<String> testCaseDropdown;
//        private JTextField associateOIDField;
//        private JTextField workAgreementIDField;
//        private JComboBox<String> timezoneDropdown;
//        private JTextField calculationTypeCodeField;
//        private JTextField calculationTypeNameField;
//        private JTextField legalEntityIDField;
//        private JTextField countrycodeField;
//        private JTextField jurisdictionIDField1;
//        private JTextField jurisdictionCodeField1;
//        private JTextField jurisdictionNameField1;
//        private JTextField jurisdictionIDField2;
//        private JTextField jurisdictionCodeField2;
//        private JTextField jurisdictionNameField2;
//        private JTextField timePeriodIDField;
//        private JTextField startDateTimeField;
//        private JTextField endDateTimeField;
//        private JTextField inTypeCodeField;
//        private JTextField outTypeCodeField;
//        private JTextField attestationsField;
//        private JTextField breaksField;
//        private JTextField shiftGroupIDField;
//        private JTextField reportDateField;
//        private JButton submitButton;
//
//        public JsonGenerator() {
//            super("JSON Generator");
//            setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//            setLayout(new BorderLayout());
//
//            JTabbedPane tabbedPane = new JTabbedPane();
//
//            // Main tab
//            JPanel mainPanel = new JPanel(new GridLayout(0, 2, 10, 10));
//
//            // Initialize dropdowns
//            String[] modules = {"Module 1", "Module 2"};
//            moduleDropdown = new JComboBox<>(modules);
//            moduleDropdown.addActionListener(this);
//            mainPanel.add(new JLabel("Select Module:"));
//            mainPanel.add(moduleDropdown);
//
//            String[] testCases = generateTestCases();
//            testCaseDropdown = new JComboBox<>(testCases);
//            testCaseDropdown.setEnabled(false); // Disable initially
//            mainPanel.add(new JLabel("Select Test Case Number:"));
//            mainPanel.add(testCaseDropdown);
//
//            // Generate associate OID
//            String associateOID = generateAssociateOID();
//            associateOIDField = new JTextField(associateOID);
//            associateOIDField.setEditable(false); // Not editable
//            mainPanel.add(new JLabel("Associate OID:"));
//            mainPanel.add(associateOIDField);
//
//            // Generate work agreement ID
//            String workAgreementID = generateWorkAgreementID();
//            workAgreementIDField = new JTextField(workAgreementID);
//            workAgreementIDField.setEditable(false); // Not editable
//            mainPanel.add(new JLabel("Work Agreement ID:"));
//            mainPanel.add(workAgreementIDField);
//
//            // Add other UI components for input fields
//            String[] timezones = {"America/New_York", "America/Los_Angeles", "America/Chicago", "America/Denver"};
//            timezoneDropdown = new JComboBox<>(timezones);
//            mainPanel.add(new JLabel("Timezone:"));
//            mainPanel.add(timezoneDropdown);
//
//            // Calculation Type Code section
//            JPanel calculationTypePanel = new JPanel(new GridLayout(0, 2));
//            calculationTypePanel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Calculation Type Code"));
//            calculationTypeCodeField = new JTextField();
//            calculationTypeNameField = new JTextField();
//            calculationTypePanel.add(new JLabel("Code:"));
//            calculationTypePanel.add(calculationTypeCodeField);
//            calculationTypePanel.add(new JLabel("Name:"));
//            calculationTypePanel.add(calculationTypeNameField);
//            mainPanel.add(calculationTypePanel);
//
//            // Legal Entity Reference section
//            JPanel legalEntityPanel = new JPanel(new GridLayout(0, 2));
//            legalEntityPanel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Legal Entity Reference"));
//            legalEntityIDField = new JTextField();
//            countrycodeField = new JTextField();
//            legalEntityPanel.add(new JLabel("Legal Entity ID:"));
//            legalEntityPanel.add(legalEntityIDField);
//            legalEntityPanel.add(new JLabel("Country Code:"));
//            legalEntityPanel.add(countrycodeField);
//            mainPanel.add(legalEntityPanel);
//
//            // Primary Worked In Jurisdictions section
//            JPanel primaryJurisdictionsPanel = new JPanel(new GridLayout(0, 2));
//            primaryJurisdictionsPanel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Primary Worked In Jurisdictions"));
//
//            // First section
//            JPanel jurisdictionSection1 = new JPanel(new GridLayout(0, 2));
//            jurisdictionIDField1 = new JTextField();
//            jurisdictionCodeField1 = new JTextField("STATE");
//            jurisdictionCodeField1.setEditable(false);
//            jurisdictionNameField1 = new JTextField();
//            jurisdictionSection1.add(new JLabel("Jurisdiction ID:"));
//            jurisdictionSection1.add(jurisdictionIDField1);
//            jurisdictionSection1.add(new JLabel("Jurisdiction Level Code:"));
//            jurisdictionSection1.add(new JLabel()); // Subheading
//            jurisdictionSection1.add(new JLabel("Code:"));
//            jurisdictionSection1.add(jurisdictionCodeField1);
//            jurisdictionSection1.add(new JLabel("Name:"));
//            jurisdictionSection1.add(jurisdictionNameField1);
//            primaryJurisdictionsPanel.add(jurisdictionSection1);
//
//            // Second section
//            JPanel jurisdictionSection2 = new JPanel(new GridLayout(0, 2));
//            jurisdictionIDField2 = new JTextField();
//            jurisdictionCodeField2 = new JTextField("FEDERAL");
//            jurisdictionCodeField2.setEditable(false);
//            jurisdictionNameField2 = new JTextField();
//            jurisdictionSection2.add(new JLabel("Jurisdiction ID:"));
//            jurisdictionSection2.add(jurisdictionIDField2);
//            jurisdictionSection2.add(new JLabel("Jurisdiction Level Code:"));
//            jurisdictionSection2.add(new JLabel()); // Subheading
//            jurisdictionSection2.add(new JLabel("Code:"));
//            jurisdictionSection2.add(jurisdictionCodeField2);
//            jurisdictionSection2.add(new JLabel("Name:"));
//            jurisdictionSection2.add(jurisdictionNameField2);
//            primaryJurisdictionsPanel.add(jurisdictionSection2);
//
//            mainPanel.add(primaryJurisdictionsPanel);
//
//            // Primary Lived In Jurisdictions section
//            JPanel livedJurisdictionsPanel = new JPanel(new GridLayout(0, 2));
//            livedJurisdictionsPanel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Primary Lived In Jurisdictions"));
//
//            // Empty sections
//            JPanel livedJurisdictionSection1 = new JPanel(new GridLayout(0, 2));
//            livedJurisdictionsPanel.add(livedJurisdictionSection1);
//
//            JPanel livedJurisdictionSection2 = new JPanel(new GridLayout(0, 2));
//            livedJurisdictionsPanel.add(livedJurisdictionSection2);
//
//            mainPanel.add(livedJurisdictionsPanel);
//
//            tabbedPane.addTab("Main", mainPanel);
//
//            // TimePeriods tab
//            JPanel timePeriodsPanel = new JPanel(new GridLayout(0, 2, 10, 10));
//            timePeriodsPanel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Time Periods"));
//
//            // Generate time period ID
//            String timePeriodID = generateTimePeriodID();
//            timePeriodIDField = new JTextField(timePeriodID);
//            timePeriodIDField.setEditable(false);
//            timePeriodsPanel.add(new JLabel("Time Period ID:"));
//            timePeriodsPanel.add(timePeriodIDField);
//
//            // Add other UI components for time period fields
//            startDateTimeField = new JTextField();
//            endDateTimeField = new JTextField();
//            inTypeCodeField = new JTextField();
//            outTypeCodeField = new JTextField();
//            attestationsField = new JTextField();
//            breaksField = new JTextField();
//            shiftGroupIDField = new JTextField();
//            reportDateField = new JTextField();
//
//            timePeriodsPanel.add(new JLabel("Start DateTime:"));
//            timePeriodsPanel.add(startDateTimeField);
//            timePeriodsPanel.add(new JLabel("End DateTime:"));
//            timePeriodsPanel.add(endDateTimeField);
//            timePeriodsPanel.add(new JLabel("In Type Code:"));
//            timePeriodsPanel.add(inTypeCodeField);
//            timePeriodsPanel.add(new JLabel("Out Type Code:"));
//            timePeriodsPanel.add(outTypeCodeField);
//            timePeriodsPanel.add(new JLabel("Attestations:"));
//            timePeriodsPanel.add(attestationsField);
//            timePeriodsPanel.add(new JLabel("Breaks:"));
//            timePeriodsPanel.add(breaksField);
//            timePeriodsPanel.add(new JLabel("Shift Group ID:"));
//            timePeriodsPanel.add(shiftGroupIDField);
//            timePeriodsPanel.add(new JLabel("Report Date:"));
//            timePeriodsPanel.add(reportDateField);
//
//            // Add submit button
//            submitButton = new JButton("Submit");
//            submitButton.addActionListener(this);
//            timePeriodsPanel.add(submitButton);
//
//            tabbedPane.addTab("Time Periods", timePeriodsPanel);
//
//            add(tabbedPane, BorderLayout.CENTER);
//
//            pack();
//            setLocationRelativeTo(null); // Center the window
//            setVisible(true);
//        }
//
//        public void actionPerformed(ActionEvent e) {
//            if (e.getSource() == moduleDropdown) {
//                testCaseDropdown.setEnabled(true);
//            } else if (e.getSource() == submitButton) {
//                saveFormDataToJson();
//                saveFormDataToExcel();
//            }
//        }
//
//        private String[] generateTestCases() {
//            String[] testCases = new String[99];
//            for (int i = 0; i < 99; i++) {
//                testCases[i] = "TC" + (i + 1);
//            }
//            return testCases;
//        }
//
//        private String generateAssociateOID() {
//            Random random = new Random();
//            StringBuilder oid = new StringBuilder("G");
//            for (int i = 0; i < 15; i++) {
//                oid.append((char) ('A' + random.nextInt(26)));
//            }
//            return oid.toString();
//        }
//
//        private String generateWorkAgreementID() {
//            Random random = new Random();
//            StringBuilder workAgreementID = new StringBuilder();
//            for (int i = 0; i < 9; i++) {
//                workAgreementID.append(random.nextInt(10));
//            }
//            workAgreementID.append("N");
//            return workAgreementID.toString();
//        }
//
//        private String generateTimePeriodID() {
//            return UUID.randomUUID().toString();
//        }
//
//        private void saveFormDataToJson() {
//            try {
//                JSONObject jsonObject = new JSONObject();
//                jsonObject.put("module", moduleDropdown.getSelectedItem().toString());
//                jsonObject.put("testCaseNumber", testCaseDropdown.getSelectedItem().toString());
//                jsonObject.put("associateOID", associateOIDField.getText());
//                jsonObject.put("workAgreementID", workAgreementIDField.getText());
//                jsonObject.put("timezone", timezoneDropdown.getSelectedItem().toString());
//
//                JSONObject calculationTypeCode = new JSONObject();
//                calculationTypeCode.put("code", calculationTypeCodeField.getText());
//                calculationTypeCode.put("name", calculationTypeNameField.getText());
//                jsonObject.put("calculationTypeCode", calculationTypeCode);
//
//                JSONObject legalEntityReference = new JSONObject();
//                legalEntityReference.put("legalEntityID", legalEntityIDField.getText());
//                legalEntityReference.put("countryCode", countrycodeField.getText());
//                jsonObject.put("legalEntityReference", legalEntityReference);
//
//                JSONObject primaryWorkedInJurisdiction1 = new JSONObject();
//                primaryWorkedInJurisdiction1.put("jurisdictionID", jurisdictionIDField1.getText());
//                primaryWorkedInJurisdiction1.put("jurisdictionLevelCode", new JSONObject().put("code", jurisdictionCodeField1.getText()).put("name", jurisdictionNameField1.getText()));
//
//                JSONObject primaryWorkedInJurisdiction2 = new JSONObject();
//                primaryWorkedInJurisdiction2.put("jurisdictionID", jurisdictionIDField2.getText());
//                primaryWorkedInJurisdiction2.put("jurisdictionLevelCode", new JSONObject().put("code", jurisdictionCodeField2.getText()).put("name", jurisdictionNameField2.getText()));
//
//                jsonObject.put("primaryWorkedInJurisdictions", new JSONObject[]{primaryWorkedInJurisdiction1, primaryWorkedInJurisdiction2});
//
//                // Add empty primaryLivedInJurisdictions
//                jsonObject.put("primaryLivedInJurisdictions", new JSONObject[0]);
//
//                // Time Periods
//                JSONObject timePeriods = new JSONObject();
//                timePeriods.put("timePeriodID", timePeriodIDField.getText());
//                timePeriods.put("startDateTime", startDateTimeField.getText());
//                timePeriods.put("endDateTime", endDateTimeField.getText());
//                timePeriods.put("inTypeCode", inTypeCodeField.getText());
//                timePeriods.put("outTypeCode", outTypeCodeField.getText());
//                timePeriods.put("attestations", attestationsField.getText());
//                timePeriods.put("breaks", breaksField.getText());
//                timePeriods.put("shiftGroupID", shiftGroupIDField.getText());
//                timePeriods.put("reportDate", reportDateField.getText());
//                jsonObject.put("timePeriods", timePeriods);
//
//                // Save to file
//                try (FileWriter file = new FileWriter("data.json")) {
//                    file.write(jsonObject.toString(4)); // Indent with 4 spaces for readability
//                }
//
//                JOptionPane.showMessageDialog(this, "Data saved to data.json");
//            } catch (IOException ex) {
//                ex.printStackTrace();
//                JOptionPane.showMessageDialog(this, "Error saving data to file", "Error", JOptionPane.ERROR_MESSAGE);
//            } catch (JSONException e) {
//                throw new RuntimeException(e);
//            }
//        }
//
//        private void saveFormDataToExcel() {
//            Workbook workbook = new XSSFWorkbook();
//            Sheet sheet = workbook.createSheet("Data");
//
//            // Create headers
//            Row headerRow = sheet.createRow(0);
//            headerRow.createCell(0).setCellValue("Module");
//            headerRow.createCell(1).setCellValue("Test Case Number");
//            headerRow.createCell(2).setCellValue("Associate OID");
//            headerRow.createCell(3).setCellValue("Work Agreement ID");
//            headerRow.createCell(4).setCellValue("Timezone");
//            headerRow.createCell(5).setCellValue("Calculation Type Code - Code");
//            headerRow.createCell(6).setCellValue("Calculation Type Code - Name");
//            headerRow.createCell(7).setCellValue("Legal Entity ID");
//            headerRow.createCell(8).setCellValue("Country Code");
//            headerRow.createCell(9).setCellValue("Primary Worked In Jurisdictions - Jurisdiction ID 1");
//            headerRow.createCell(10).setCellValue("Primary Worked In Jurisdictions - Jurisdiction Level Code 1");
//            headerRow.createCell(11).setCellValue("Primary Worked In Jurisdictions - Jurisdiction Name 1");
//            headerRow.createCell(12).setCellValue("Primary Worked In Jurisdictions - Jurisdiction ID 2");
//            headerRow.createCell(13).setCellValue("Primary Worked In Jurisdictions - Jurisdiction Level Code 2");
//            headerRow.createCell(14).setCellValue("Primary Worked In Jurisdictions - Jurisdiction Name 2");
//            headerRow.createCell(15).setCellValue("Time Period ID");
//            headerRow.createCell(16).setCellValue("Start DateTime");
//            headerRow.createCell(17).setCellValue("End DateTime");
//            headerRow.createCell(18).setCellValue("In Type Code");
//            headerRow.createCell(19).setCellValue("Out Type Code");
//            headerRow.createCell(20).setCellValue("Attestations");
//            headerRow.createCell(21).setCellValue("Breaks");
//            headerRow.createCell(22).setCellValue("Shift Group ID");
//            headerRow.createCell(23).setCellValue("Report Date");
//
//            // Add data
//            Row dataRow = sheet.createRow(1);
//            dataRow.createCell(0).setCellValue(moduleDropdown.getSelectedItem().toString());
//            dataRow.createCell(1).setCellValue(testCaseDropdown.getSelectedItem().toString());
//            dataRow.createCell(2).setCellValue(associateOIDField.getText());
//            dataRow.createCell(3).setCellValue(workAgreementIDField.getText());
//            dataRow.createCell(4).setCellValue(timezoneDropdown.getSelectedItem().toString());
//            dataRow.createCell(5).setCellValue(calculationTypeCodeField.getText());
//            dataRow.createCell(6).setCellValue(calculationTypeNameField.getText());
//            dataRow.createCell(7).setCellValue(legalEntityIDField.getText());
//            dataRow.createCell(8).setCellValue(countrycodeField.getText());
//            dataRow.createCell(9).setCellValue(jurisdictionIDField1.getText());
//            dataRow.createCell(10).setCellValue(jurisdictionCodeField1.getText());
//            dataRow.createCell(11).setCellValue(jurisdictionNameField1.getText());
//            dataRow.createCell(12).setCellValue(jurisdictionIDField2.getText());
//            dataRow.createCell(13).setCellValue(jurisdictionCodeField2.getText());
//            dataRow.createCell(14).setCellValue(jurisdictionNameField2.getText());
//            dataRow.createCell(15).setCellValue(timePeriodIDField.getText());
//            dataRow.createCell(16).setCellValue(startDateTimeField.getText());
//            dataRow.createCell(17).setCellValue(endDateTimeField.getText());
//            dataRow.createCell(18).setCellValue(inTypeCodeField.getText());
//            dataRow.createCell(19).setCellValue(outTypeCodeField.getText());
//            dataRow.createCell(20).setCellValue(attestationsField.getText());
//            dataRow.createCell(21).setCellValue(breaksField.getText());
//            dataRow.createCell(22).setCellValue(shiftGroupIDField.getText());
//            dataRow.createCell(23).setCellValue(reportDateField.getText());
//
//            // Save to file
//            try (FileOutputStream fileOut = new FileOutputStream("data.xlsx")) {
//                workbook.write(fileOut);
//                workbook.close();
//                JOptionPane.showMessageDialog(this, "Data saved to data.xlsx");
//            } catch (IOException ex) {
//                ex.printStackTrace();
//                JOptionPane.showMessageDialog(this, "Error saving data to Excel file", "Error", JOptionPane.ERROR_MESSAGE);
//            }
//        }
//
//        public static void main(String[] args) {
//            SwingUtilities.invokeLater(() -> new JsonGenerator());
//        }
//
//}
