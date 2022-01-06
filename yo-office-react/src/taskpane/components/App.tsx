import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { tryCatch } from "../lib/utils";

/* global console, Excel, require  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default function App(appProps:AppProps){
const [appState, setAppState] = React.useState([])

React.useEffect(()=>{
  setAppState([ {
    icon: "Ribbon",
    primaryText: "Achieve more with Office integration",
  },
  {
    icon: "Unlock",
    primaryText: "Unlock features and functionality",
  },
  {
    icon: "Design",
    primaryText: "Create and visualize like a pro",
  },
  {
    icon: "OfficeLogo",
    primaryText: "A test for the add in",
  },])
},[])

async function handleClick(){  
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
};

function createTable() {
  Excel.run(function (context) {

      // TODO1: Queue table creation logic here.
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
expensesTable.name = "ExpensesTable";

      // TODO2: Queue commands to populate the table with data.
      expensesTable.getHeaderRowRange().values =
    [["Date", "Merchant", "Category", "Amount"]];

expensesTable.rows.add(null /*add at the end*/, [
    ["1/1/2017", "The Phone Company", "Communications", "120"],
    ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
    ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
    ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
    ["1/11/2017", "Bellows College", "Education", "350.1"],
    ["1/15/2017", "Trey Research", "Other", "135"],
    ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
]);

      // TODO3: Queue commands to format the table.
      expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
expensesTable.getRange().format.autofitColumns();
expensesTable.getRange().format.autofitRows();

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function filterTable() {
  Excel.run(function (context) {

      // TODO1: Queue commands to filter out all expense categories except Groceries and Education.
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
var categoryFilter = expensesTable.columns.getItem('Category').filter;
categoryFilter.applyValuesFilter(['Education', 'Groceries']);

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function sortTable() {
  Excel.run(function (context) {

      // TODO1: Queue commands to sort the table by Merchant name.
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
var sortFields = [
    {
        key: 1,            // Merchant column
        ascending: false,
    }
];

expensesTable.sort.apply(sortFields);

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function createChart() {
  Excel.run(function (context) {

      // TODO1: Queue commands to get the range of data to be charted.
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
var dataRange = expensesTable.getDataBodyRange();

      // TODO2: Queue command to create the chart and define its type.
      var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'Auto');

      // TODO3: Queue commands to position and format the chart.
      chart.setPosition("A15", "F30");
      chart.title.text = "Expenses";
      chart.legend.position = "Right";
      chart.legend.format.fill.setSolidColor("white");
      chart.dataLabels.format.font.size = 15;
      chart.dataLabels.format.font.color = "black";
      chart.series.getItemAt(0).name = 'Value in \u20AC';

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function freezeHeader() {
  Excel.run(function (context) {

      // TODO1: Queue commands to keep the header visible when the user scrolls.
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
currentWorksheet.freezePanes.freezeRows(1);

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

const { title, isOfficeInitialized } = appProps;

if (!isOfficeInitialized) {
  return (
    <Progress
      title={title}
      logo={require("./../../../assets/logo-filled.png")}
      message="Please sideload your addin to see app body."
    />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo={require("./../../../assets/logo-filled.png")} title={appProps.title} message="Welcome" />
      <HeroList message="Discover what Office Add-ins can do for you today!" items={appState}>
        <p className="ms-font-l">
          Modify the source files, then click <b>Run</b>.
        </p>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={handleClick}>
          Run
        </DefaultButton>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={createTable}>
          Create Table
        </DefaultButton>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={filterTable}>
          Filter Table
        </DefaultButton>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={sortTable}>
          Sort Table
        </DefaultButton>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={createChart}>
          Create Chart
        </DefaultButton>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={freezeHeader}>
          Freeze Header
        </DefaultButton>
      </HeroList>
    </div>
  );
}