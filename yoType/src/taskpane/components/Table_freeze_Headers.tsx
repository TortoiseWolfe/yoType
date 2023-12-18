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

const FreezeTableHeader: React.FC = () => {
  const styles = useStyles();

  const freezeHeader = async () => {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      // Freeze the top row of the active worksheet
      sheet.freezePanes.freezeRows(1);
      await context.sync();
    });
  };

  return (
    <div className={styles.buttonContainer}>
      <Button appearance="primary" size="large" onClick={freezeHeader}>
        Freeze Table Header in Excel
      </Button>
    </div>
  );
};

export default FreezeTableHeader;
