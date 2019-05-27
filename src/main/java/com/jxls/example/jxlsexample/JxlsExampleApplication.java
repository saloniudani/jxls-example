package com.jxls.example.jxlsexample;

import com.jxls.example.jxlsexample.model.Employee;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.util.*;

@SpringBootApplication
public class JxlsExampleApplication {

    public static void main(String[] args) throws ParseException, IOException {

        SpringApplication.run(JxlsExampleApplication.class, args);

        List<Employee> employees = generateSampleEmployeeData();
        executeGridTemplateDemo(employees);
        executeCardTemplateDemo(employees);
        executeCardPerSheetTemplateDemo(employees);
    }

    private static void executeGridTemplateDemo(List<Employee> employees) throws IOException {
        Long now = Instant.now().getEpochSecond();
        try(InputStream is = JxlsExampleApplication.class.getResourceAsStream("grid_template.xls")) {
            try(OutputStream os = new FileOutputStream("./grid_output_"+String.valueOf(now)+".xls")) {
                Context context = new Context();
                context.putVar("headers", Arrays.asList("Name", "Birthday", "Payment","Bonus"));
                context.putVar("data", employees);
                context.putVar("projectName","jxls-example");
                context.putVar("createdOn",new Date().toString());
                context.putVar("author","Name Surname");
                JxlsHelper.getInstance().processGridTemplateAtCell(is, os, context, "name,birthDate,payment,bonus", "Sheet1!A1");
            }
        }
    }

    private static void executeCardTemplateDemo(List<Employee> employees) throws IOException {
        Long now = Instant.now().getEpochSecond();
        try(InputStream is = JxlsExampleApplication.class.getResourceAsStream("card_template.xls")) {
            try(OutputStream os = new FileOutputStream("./card_output_"+String.valueOf(now)+".xls")) {
                Context context = new Context();
                context.putVar("employees", employees);
                context.putVar("projectName","jxls-example");
                context.putVar("createdOn",new Date().toString());
                context.putVar("author","Name Surname");
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }

    private static void executeCardPerSheetTemplateDemo(List<Employee> employees) throws IOException {
        Long now = Instant.now().getEpochSecond();
        try(InputStream is = JxlsExampleApplication.class.getResourceAsStream("card_per_sheet_template.xls")) {
            try(OutputStream os = new FileOutputStream("./card_per_sheet_output_"+String.valueOf(now)+".xls")) {
                Context context = new Context();
                context.putVar("employees", employees);
                context.putVar("projectName","jxls-example");
                context.putVar("createdOn",new Date().toString());
                context.putVar("author","Name Surname");
                JxlsHelper.getInstance().setUseFastFormulaProcessor(false).processTemplate(is, os, context);
            }
        }
    }

    private static List<Employee> generateSampleEmployeeData() throws ParseException {
        List<Employee> employees = new ArrayList<>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        employees.add(Employee.builder().name("Elsa").birthDate("1970-Jul-10").payment(1500l).bonus(0.15f).build());
        employees.add(Employee.builder().name("Oleg").birthDate("1970-Jul-10").payment(2300l).bonus(0.15f).build());
        employees.add(Employee.builder().name("Neil").birthDate("1975-Oct-05").payment(2500l).bonus(0.25f).build());
        return employees;
    }

}
