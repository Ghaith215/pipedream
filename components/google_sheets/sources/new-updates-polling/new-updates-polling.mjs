import googleSheets from "../../google_sheets.app.mjs";
import common from "../common/new-updates.mjs";
import base from "../common/http-based/base.mjs";
import { DEFAULT_POLLING_SOURCE_TIMER_INTERVAL } from "@pipedream/platform";

export default {
  ...common,
  key: "google_sheets-new-updates-polling",
  name: "New Updates",
  description: "Emit new event each time a row or cell is updated in a spreadsheet.",
  version: "0.0.1",
  dedupe: "unique",
  type: "source",
  props: {
    googleSheets,
    db: "$.service.db",
    timer: {
      type: "$.interface.timer",
      static: {
        intervalSeconds: DEFAULT_POLLING_SOURCE_TIMER_INTERVAL,
      },
    },
    // eslint-disable-next-line pipedream/props-label, pipedream/props-description
    alert: {
      type: "alert",
      alertType: "info",
      content: "**Important**: If your sheet has more than 1000 rows, please set the **Monitoring Range** below to avoid performance issues and potential disruptions. Example: `A1:Z1000` to monitor the first 1000 rows with columns A through Z.",
    },
    watchedDrive: {
      propDefinition: [
        googleSheets,
        "watchedDrive",
      ],
      description: "Defaults to My Drive. To select a [Shared Drive](https://support.google.com/a/users/answer/9310351) instead, select it from this list.",
    },
    sheetID: {
      propDefinition: [
        googleSheets,
        "sheetID",
        (c) => ({
          driveId: googleSheets.methods.getDriveId(c.watchedDrive),
        }),
      ],
    },
    worksheetIDs: {
      propDefinition: [
        googleSheets,
        "worksheetIDs",
        (c) => ({
          sheetId: c.sheetID,
        }),
      ],
      type: "integer[]",
      label: "Worksheet ID(s)",
      description: "Select one or more worksheet(s), or provide an array of worksheet IDs.",
    },
    monitoringRange: {
      type: "string",
      label: "Monitoring Range",
      description: "The A1 notation range to monitor for changes (e.g., `A1:B100` or `Sheet1!A1:Z1000`). If not specified, the entire sheet will be monitored up to 10000 rows. **Recommended for sheets with more than 1000 rows**.",
      optional: true,
    },
  },
  methods: {
    ...base.methods,
    ...common.methods,
    getMonitoringRange() {
      return this.monitoringRange;
    },
    async getContentDiff(spreadsheet, worksheet) {
      const sheetId = this.getSheetId();
      const baseId = `${spreadsheet.spreadsheetId}${worksheet.properties.sheetId}`;
      const oldValues = this._getBatchedSheetValues(baseId) || null;

      // Use monitoring range if specified, otherwise use worksheet title
      const range = this.monitoringRange
        ? this.monitoringRange
        : worksheet.properties.title;

      const currentValues = await this.googleSheets.getSpreadsheetValues(
        sheetId,
        range,
      );
      return {
        oldValues,
        currentValues,
      };
    },
  },
  hooks: {
    async deploy() {
      await this.takeSheetSnapshot();
    },
  },
  async run() {
    const spreadsheet = await this.googleSheets.getSpreadsheet(this.sheetID);
    return this.processSpreadsheet(spreadsheet);
  },
};
