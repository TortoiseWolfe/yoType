import * as React from "react";
import { Button, makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  buttonContainer: {
    marginTop: "20px",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
});

const FilterTable: React.FC = () => {
  const styles = useStyles();

  const applyFilter = async () => {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const table = sheet.tables.getItem("ExpensesTable");
      const categoryColumn = table.columns.getItem("Category");
      categoryColumn.filter.applyValuesFilter(["Education", "Groceries"]);
      await context.sync();
    });
  };

  return (
    <div className={styles.buttonContainer}>
      <Button appearance="primary" size="large" onClick={applyFilter}>
        Apply Filter in Excel
      </Button>
    </div>
  );
};

export default FilterTable;
