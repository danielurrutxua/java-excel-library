# OpenXLSX
OpenXLSX is a robust Java library for easy and efficient Excel file manipulation.

## Features
- Facilitates the creation of new Excel (.xlsx) files with a straightforward Java API.
- Specializes in writing data to Excel files, enabling quick and easy generation of sheets and tables.
- Leverages the power of Apache POI to allow seamless Excel operations.
- Main advantage lies in its ability to create structured tables from a list of class instances, streamlining the process of transferring data into an Excel format.
- Currently supports writing and creating Excel files, with read and modify functionalities planned for future updates.
- Compatible with Java 8 and newer.

## Installation

### Gradle
Since the library is hosted on GitHub Packages, you'll need to add GitHub Packages as a repository in your `build.gradle`:

```groovy
repositories {
    mavenCentral()
    maven { url 'https://maven.pkg.github.com/danielurrutxua/java-excel-library' }
    credentials {
                username = <YOUR_GITHUB_USERNAME>
                password = <YOUR_GITHUB_PASSWORD>
    }

}
```

Then add the dependency:

```groovy
dependencies {
    implementation 'com.github.danielurrutxua:open-xlsx:1.0.1'
}
```

### Maven

To include OpenXLSX in your Maven project, you will need to add the GitHub package repository to your `pom.xml` and then include the dependency.

Add the following repository to your `pom.xml` file:

```xml
<repositories>
    <repository>
        <id>github-openxlsx</id>
        <url>https://maven.pkg.github.com/danielurrutxua/java-excel-library</url>
    </repository>
</repositories>
```
Then add the following dependency to your `pom.xml`:

```xml
<dependencies>
    <dependency>
        <groupId>com.github.danielurrutxua</groupId>
        <artifactId>open-xlsx</artifactId>
        <version>1.0.1</version>
    </dependency>
</dependencies>
```
## Usage

OpenXLSX makes it simple to create Excel files and manipulate their contents. Here's a quick start guide:

### Creating a New Sheet and Adding Data

```java
import com.github.danielurrutxua.openxlsx.OpenXLSX;
import java.util.List;

public class ExcelUsageExample {
    public static void main(String[] args) {
        // Initialize an Excel Builder
        OpenXLSX excelBuilder = new OpenXLSX();

        // Initialize workbook
        excelbuilder.initWorkbook("summary_2023")

        // Create a new sheet with a name
        excelBuilder.createSheet("Employee Data");

        // List of any class
        List<Employee> employees = new ArrayList<>();
        employeeList.add(new Employee(1, "John", "Doe", "txomin.doe@example.com"));
        employeeList.add(new Employee(2, "Jane", "Doe", "kaizka.doe@example.com"));
        employeeList.add(new Employee(3, "Alice", "Smith", "maialen.smith@example.com"));
        employeeList.add(new Employee(4, "Bob", "Brown", "perro.sanchez@example.com"));
        employeeList.add(new Employee(5, "Charlie", "Davis", "charlie.davis@example.com"));


        // Add data to the sheet
        excelBuilder.setData(employees);

        // Generate the Excel file
        File excelFile = excelBuilder.generateOutputFile("EmployeeData.xlsx");

        // ... now you can write the file to disk, send it over the network, etc.
    }
}
```

### Defining Data Model with Annotations

The `Employee` class with `@OpenXLSXColumn` annotations might look like this:

```java
import com.github.danielurrutxua.openxlsx.annotations.OpenXLSXColumn;

public class Employee {
    @OpenXLSXColumn(name = "ID")
    private long id;

    @OpenXLSXColumn(name = "First Name")
    private String firstName;

    @OpenXLSXColumn(name = "Last Name")
    private String lastName;

    @OpenXLSXColumn(name = "Department")
    private String department;

    // Standard constructors, getters and setters below...
}
```

### Example
```java
import com.github.danielurrutxua.openxlsx.annotations.OpenXLSXColumn;

public class Employee {
    @OpenXLSXColumn(name = "ID")
    private long id;

    @OpenXLSXColumn(name = "First Name")
    private String firstName;

    @OpenXLSXColumn(name = "Last Name")
    private String lastName;

    @OpenXLSXColumn(name = "Department")
    private String department;

    // Standard constructors, getters and setters below...
}
```
