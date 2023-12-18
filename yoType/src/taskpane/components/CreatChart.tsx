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

const CreateChart: React.FC = () => {
  const styles = useStyles();

  const createChart = async () => {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = sheet.tables.getItem('ExpensesTable');
      const dataRange = expensesTable.getDataBodyRange();
      const chart = sheet.charts.add('ColumnClustered', dataRange, 'Auto');

      chart.setPosition("A15", "F30");
      chart.title.text = "Expenses";
      chart.legend.position = "Right";
      chart.legend.format.fill.setSolidColor("white");
      chart.dataLabels.format.font.size = 15;
      chart.dataLabels.format.font.color = "black";
      chart.series.getItemAt(0).name = 'Value in €';

      await context.sync();
    });
  };

  return (
    <div className={styles.buttonContainer}>
      <Button appearance="primary" size="large" onClick={createChart}>
        Create Chart in Excel
      </Button>
    </div>
  );
};

export default CreateChart;
