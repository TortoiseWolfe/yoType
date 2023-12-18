import * as React from "react";
import { useState } from "react";
import { Button, makeStyles } from "@fluentui/react-components";
//import { Excel } from "../office-document"; // Assuming office-document module handles Office.js interactions

const useStyles = makeStyles({
  buttonContainer: {
    marginTop: "20px",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
});

const TableInsertion: React.FC = () => {
  const createTable = async () => {
    try {
      await Excel.run(async (context) => {
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        const expensesTable = currentWorksheet.tables.add("A1:D1", true);
        expensesTable.name = "ExpensesTable";
        expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];
        expensesTable.rows.add(null, [
          ["1/1/2017", "The Phone Company", "Communications", "120"],
          ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
          ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
          ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
          ["1/11/2017", "Bellows College", "Education", "350.1"],
          ["1/15/2017", "Trey Research", "Other", "135"],
          ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
        ]);
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['"$"#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();

        await context.sync();
      });
    } catch (error) {
      console.error("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    }
  };

  const styles = useStyles();

  return (
    <div className={styles.buttonContainer}>
      <Button appearance="primary" size="large" onClick={createTable}>
        Create Table in Excel
      </Button>
    </div>
  );
};

export default TableInsertion;
