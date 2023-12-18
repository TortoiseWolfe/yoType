import * as React from "react";
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

const SortTable: React.FC = () => {
  const applySort = async () => {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const table = sheet.tables.getItem("ExpensesTable");
      const sortField = {
        key: 1,            // Index of the "Merchant" column
        ascending: false   // Descending order
      };
      table.sort.apply([sortField]);
      await context.sync();
    });
  };

const styles = useStyles();

  return (
    <div className={styles.buttonContainer}>
    <Button appearance="primary" size="large" onClick={applySort}>
      Sort Table in Excel
    </Button>
  </div>
  );
};

export default SortTable;
