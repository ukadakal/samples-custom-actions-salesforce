# SpreadsheetWeb Custom Actions Example - Salesforce integration

SpreadsheetWeb is a platform for building and generating web applications and complex calculation engines from Excel models. The software is produced and distributed by Pagos Inc.

This repository contains sample code that aims to demonstrate how to inject custom code that can communicate with the [Salesforce](https://www.salesforce.com) API. The provided sample code submits a record to the Salesforce data structure, but the same basic principles can be applied to any functions permitted by the target service.

> The sample can be used only with SpreadsheetWeb Server Packages - must have an active [server license](https://www.spreadsheetweb.com/server-pricing/).

> You will also need to have an active Salesforce account in order to request access to their web service. You can register for a free trial account if you are not a paid subscriber.

## Download

If you have `git` command line installed on your computer, you can download the sample by running:

```bash
> git clone https://github.com/spreadsheetweb/samples-custom-actions-salesforce
```

Alternatively, you can click the `Clone or Download` button on the GitHub page.

## How do Custom Actions work in the SpreadsheetWeb platform?

- Custom actions must be written as a .NET Framework Library. In our sample, we have generated a solution in C#, which can be viewed and modified with Microsoft Visual Studio or other compatible software.
- When creating a new library project you must add references to several prerequisite .NET libraries that will be provided to you. Alternatively, you can utilize the ones that are included in the sample or contact our support for more information on versions that are compatible with your server version.
- The copy included in the sample can be found under the _Pagos References_ folder. These are also added as references to the Visual Studio project.

    ```bash
    > Pagos.Designer.Interfaces.External.dll
    > Pagos.SpreadsheetWeb.Web.Api.Objects.dll
    ```

- In the imports section of your class file, make sure to include the following namespaces from the aforementioned libraries:

    ```C#
    using Pagos.Designer.Interfaces.External.CustomHooks;
    using Pagos.Designer.Interfaces.External.Messaging;
    using Pagos.SpreadsheetWeb.Web.Api.Objects.Calculation;
    ```

- The library needs to implement one or more interfaces exposed by the `Pagos.Designer.Interfaces.External.CustomHooks` namespace. More details regarding these interfaces can be found at the following help page: [Custom Actions in Designer](https://pagosinc.atlassian.net/wiki/spaces/SSWEB/pages/501186561/Custom+Actions+in+Designer).

    - `IBeforeCalculation`
    - `IAfterCalculation`
    - `IBeforeSave`
    - `IAfterSave`
    - `IBeforePrint`
    - `IAfterPrint`
    - `IBeforeEmail`
    - `IAfterEmail`

- Each of those interfaces expose specific methods, which you will need to implement. These can be seen in the sample.

## How to run the sample

1. Download the sample. This is a Visual Studio solution that includes a single C# class library.
2. Open the solution file in Visual Studio. Before proceeding, you should set your Salesforce account credentials.

    ```C#
    namespace SalesForceSample
    {
    
        public class LeadTest : IAfterCalculation
        {
            private string accountName = "your_salesforce_account_email";
            private string accountPwd = "password";
            
            ...
        }
        
        ...
    }    
    ```
3. Compile the solution. The output of the compilation will be the **SalesForceSample.dll** library, which can subsequently be uploaded as a custom action.
4. Create a Designer application on your SpreadsheetWeb server. For this, you will need an Excel file, which will act as the primary calculation. This can be found under the _Sample_ directory. You can also review [this link](https://pagosinc.atlassian.net/wiki/spaces/SSWEB/pages/35954/Custom+Applications) for more detailed on creating an application. **Important Note:** Remember to select  _Designer_ as the application type.
5. Once the application is created, navigate to _Custom Actions_, create a **New** custom action and submit the previously compiled **SalesForceSample.dll** assembly file.
6. You will also need to create a user interface for the application. Navigate to the **User Interface Designer** and add controls to the default home page. Associate these controls with the named ranges from the Excel file. Alternatively you may use the **Generate** button, which will attempt to auto-generate a basic user interface based on the Excel calculation model's named range metadata. As video tutorial can be found at the following link: [User Interface in SpreadsheetWEB Designer](https://www.spreadsheetweb.com/project/user-interface-designer/). An example user interface can be seen below.

    ![UI-and-button-with-hook-attached.png](./Images/UI-and-button-with-hook-attached.png)
    
    > **Important Note:** In order for the custom action code to be executed, it will need to be attached to a button, as shown in the screenshot above. 
    
7. Preview the page or Publish the application.
8. Open the application and enter an email address and a name in the corresponding textboxes.
9. Click the **Submit** button.

## What is the sample about?

Upon clicking the **Submit** button, the application should connect to the Salesforce web service and generate a new [Lead](https://developer.salesforce.com/docs/atlas.en-us.sfFieldRef.meta/sfFieldRef/salesforce_field_reference_Lead.htm#!) entry, which you can subsequently browsed from the Salesforce Management Panel.


