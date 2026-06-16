/**
 * Chrome Extension Communication Service
 *
 * Centralized service for managing communication with the Student Retention Kit Chrome Extension.
 * Handles extension detection, keep-alive pings, and message relay between the add-in and extension.
 */
import { sheetNameDateVariants } from './excel-helpers.js';
import { coalesceHighlightRequests } from './highlightQueue.js';

class ChromeExtensionService {
  constructor() {
    // Connection state
    this.isExtensionInstalled = false;
    this.isChecking = false;

    // Intervals and timeouts
    this.pingInterval = null;
    this.keepAliveInterval = null;
    this.messageListenerWatchdogInterval = null;

    // Excel API health tracking
    this.lastExcelHealthCheck = null;
    this.excelHealthy = null; // null = unknown, true/false = last check result
    this.highlightRetryCount = 0;
    this.MAX_HIGHLIGHT_RETRIES = 3;
    this.EXCEL_RUN_TIMEOUT_MS = 30000; // 30 seconds

    // Highlight request queue. Incoming SRK_HIGHLIGHT_STUDENT_ROW requests
    // are enqueued and executed by a single worker so only one highlight
    // Excel.run is in flight at a time; everything that arrives while a run
    // is in progress gets coalesced into bulk operations. Without this, a
    // burst of requests (many students found by the submission checker at
    // once) runs concurrent Excel.runs that each load the whole used range,
    // which overwhelms Excel Online and drops highlights.
    this._highlightQueue = [];
    this._highlightWorkerActive = false;

    // When false, this runtime ignores highlight requests entirely. The
    // taskpane sets this to false so only the long-lived commands runtime
    // executes highlights — otherwise both runtimes receive every request
    // (the extension broadcasts to all frames) and do the same work twice.
    this.highlightExecutionEnabled = true;

    // Callbacks for extension state changes
    this.listeners = new Set();

    // Message handler bound to this instance
    this.handleMessage = this.handleMessage.bind(this);

    // Track whether message listener is attached
    this._messageListenerAttached = false;

    // Setup message listener
    this.setupMessageListener();
  }

  /**
   * Setup the global message listener for extension responses.
   * Tracks attachment state so we can detect and recover if the listener is lost.
   */
  setupMessageListener() {
    window.addEventListener("message", this.handleMessage);
    this._messageListenerAttached = true;
    console.log("ChromeExtensionService: Message listener attached.");
  }

  /**
   * Verify the message listener is still attached by posting a test message
   * to ourselves and checking if we receive it. If the listener appears dead,
   * re-attach it.
   */
  verifyMessageListener() {
    const testId = `_srk_listener_check_${Date.now()}`;
    let received = false;

    const tempHandler = (event) => {
      if (event.data && event.data._srkListenerCheck === testId) {
        received = true;
      }
    };

    window.addEventListener("message", tempHandler);
    window.postMessage({ _srkListenerCheck: testId }, "*");

    // Check after a short delay
    setTimeout(() => {
      window.removeEventListener("message", tempHandler);
      if (!received) {
        console.warn("ChromeExtensionService: Message listener self-check FAILED. Re-attaching listener.");
        this._messageListenerAttached = false;
        // Re-attach the main handler
        window.removeEventListener("message", this.handleMessage);
        window.addEventListener("message", this.handleMessage);
        this._messageListenerAttached = true;
        console.log("ChromeExtensionService: Message listener re-attached.");
      }
    }, 500);
  }

  /**
   * Start a watchdog that periodically verifies the message listener is alive.
   * This helps recover from cases where the Office Add-in iframe context
   * is silently recycled.
   * @param {number} interval - Milliseconds between checks (default: 60000 = 1 min)
   */
  startMessageListenerWatchdog(interval = 60000) {
    if (this.messageListenerWatchdogInterval) return;

    console.log("ChromeExtensionService: Starting message listener watchdog.");
    this.messageListenerWatchdogInterval = setInterval(() => {
      this.verifyMessageListener();
    }, interval);
  }

  /**
   * Stop the message listener watchdog.
   */
  stopMessageListenerWatchdog() {
    if (this.messageListenerWatchdogInterval) {
      clearInterval(this.messageListenerWatchdogInterval);
      this.messageListenerWatchdogInterval = null;
    }
  }

  /**
   * Run an Excel.run() call with a timeout. If the Excel API does not respond
   * within the timeout, the returned promise rejects with a descriptive error.
   * This helps detect stale sessions in Excel Online.
   *
   * @param {Function} batchFn - The callback to pass to Excel.run(callback)
   * @param {number} [timeoutMs] - Max milliseconds to wait (default: this.EXCEL_RUN_TIMEOUT_MS)
   * @returns {Promise<void>}
   */
  excelRunWithTimeout(batchFn, timeoutMs) {
    const timeout = timeoutMs || this.EXCEL_RUN_TIMEOUT_MS;

    return new Promise((resolve, reject) => {
      let settled = false;

      const timer = setTimeout(() => {
        if (!settled) {
          settled = true;
          const err = new Error(
            `Excel API did not respond within ${timeout / 1000}s. ` +
            "The session may have expired. Try refreshing the page."
          );
          err.code = "Timeout";
          reject(err);
        }
      }, timeout);

      Excel.run(batchFn)
        .then((result) => {
          if (!settled) {
            settled = true;
            clearTimeout(timer);
            resolve(result);
          }
        })
        .catch((err) => {
          if (!settled) {
            settled = true;
            clearTimeout(timer);
            reject(err);
          }
        });
    });
  }

  /**
   * Quick health check of the Excel API. Performs a minimal Excel.run() to
   * verify the API is responsive.
   * @param {number} [timeoutMs=5000] - How long to wait before declaring unhealthy
   * @returns {Promise<boolean>} true if healthy, false if not
   */
  async checkExcelHealth(timeoutMs = 5000) {
    if (typeof window.Excel === "undefined") {
      this.excelHealthy = false;
      return false;
    }

    try {
      await this.excelRunWithTimeout(async (context) => {
        // Minimal operation: just load the active worksheet name
        context.workbook.worksheets.getActiveWorksheet().load("name");
        await context.sync();
      }, timeoutMs);
      this.excelHealthy = true;
      this.lastExcelHealthCheck = Date.now();
      return true;
    } catch (error) {
      console.warn("ChromeExtensionService: Excel health check failed:", error.message);
      this.excelHealthy = false;
      this.lastExcelHealthCheck = Date.now();
      return false;
    }
  }

  /**
   * Determine whether an error is retryable (session/timeout issues vs
   * permanent failures like missing sheets or bad parameters).
   * @param {Error} error
   * @returns {boolean}
   */
  isRetryableError(error) {
    if (!error) return false;

    // Timeout errors are retryable
    if (error.code === "Timeout") return true;

    // InvalidObjectPath usually means the session/context is stale
    if (error.code === "InvalidObjectPath") return true;

    // GeneralException can be transient
    if (error.code === "GeneralException") return true;

    // Network errors are retryable
    if (error.message && error.message.toLowerCase().includes("network")) return true;

    // Session-related errors
    if (error.message && (
      error.message.toLowerCase().includes("session") ||
      error.message.toLowerCase().includes("expired") ||
      error.message.toLowerCase().includes("timed out") ||
      error.message.toLowerCase().includes("request failed")
    )) return true;

    return false;
  }

  /**
   * Handle incoming messages from the Chrome extension
   */
  handleMessage(event) {
    if (!event.data || !event.data.type) return;

    // Skip our own internal listener-check messages
    if (event.data._srkListenerCheck) return;

    switch (event.data.type) {
      case "SRK_EXTENSION_INSTALLED":
        console.log("ChromeExtensionService: Extension detected!");
        this.handleExtensionDetected();
        break;

      case "SRK_HIGHLIGHT_STUDENT_ROW":
        if (!this.highlightExecutionEnabled) {
          console.log("ChromeExtensionService: Highlight execution disabled in this runtime; ignoring request (the commands runtime handles it).");
          break;
        }
        console.log("ChromeExtensionService: Highlight student row request received:", event.data);
        console.log("ChromeExtensionService: Excel API available:", typeof window.Excel !== "undefined");
        console.log("ChromeExtensionService: Last Excel health:", this.excelHealthy, "checked at:", this.lastExcelHealthCheck ? new Date(this.lastExcelHealthCheck).toISOString() : "never");
        this.enqueueHighlightRequest(event.data.data);
        break;

      case "SRK_NAVIGATE_TO_STUDENT":
        console.log("ChromeExtensionService: Navigate to student request received:", event.data);
        this.handleNavigateToStudent(event.data.data);
        break;

      case "SRK_IMPORT_MASTER_LIST":
        console.log("ChromeExtensionService: Import master list request received:", event.data);
        // Forward to listeners (background-service.js will handle the actual import)
        this.notifyListeners({ type: "message", data: event.data });
        break;

      // Add more message types here as needed
      default:
        // Forward unknown messages to listeners
        this.notifyListeners({ type: "message", data: event.data });
    }
  }

  /**
   * Handle when extension is detected
   */
  handleExtensionDetected() {
    if (!this.isExtensionInstalled) {
      this.isExtensionInstalled = true;
      this.stopPinging();
      this.notifyListeners({ type: "installed", installed: true });
    }
  }

  /**
   * Send a highlight confirmation message back to the Chrome extension via window.postMessage
   * @param {string} syStudentId - The student ID (may be empty if unknown)
   * @param {"success"|"error"} status - Whether the highlight succeeded or failed
   * @param {string} message - Human-readable description of the result
   */
  sendHighlightConfirmation(syStudentId, status, message) {
    const confirmation = {
      type: "SRK_HIGHLIGHT_CONFIRMATION",
      data: {
        syStudentId: syStudentId || "",
        status,
        message,
        timestamp: Date.now()
      }
    };

    console.log(`ChromeExtensionService: Sending highlight confirmation (${status}):`, confirmation);
    this.sendMessage(confirmation);
  }

  /**
   * Enable or disable highlight execution for this runtime. Both the
   * taskpane and the commands runtime receive every highlight request (the
   * extension broadcasts to all frames); the taskpane disables execution so
   * the always-alive commands runtime is the single executor.
   * @param {boolean} enabled
   */
  setHighlightExecutionEnabled(enabled) {
    this.highlightExecutionEnabled = !!enabled;
    console.log(`ChromeExtensionService: Highlight execution ${this.highlightExecutionEnabled ? "enabled" : "disabled"} for this runtime.`);
  }

  /**
   * Queue a highlight request for the serialized worker. Requests that pile
   * up while a highlight Excel.run is in flight are coalesced into bulk
   * operations (one per target sheet/format group) before execution.
   * @param {Object} payload - SRK_HIGHLIGHT_STUDENT_ROW data (single or bulk shape)
   */
  enqueueHighlightRequest(payload) {
    this._highlightQueue.push(payload);
    this._runHighlightQueue();
  }

  /**
   * Worker loop for the highlight queue. Only one instance runs at a time;
   * each iteration drains everything queued so far, coalesces it, and
   * executes the resulting payloads sequentially.
   */
  async _runHighlightQueue() {
    if (this._highlightWorkerActive) return;
    this._highlightWorkerActive = true;

    try {
      while (this._highlightQueue.length > 0) {
        const drained = this._highlightQueue.splice(0, this._highlightQueue.length);
        const coalesced = coalesceHighlightRequests(drained);
        if (drained.length > coalesced.length) {
          console.log(`ChromeExtensionService: Coalesced ${drained.length} queued highlight requests into ${coalesced.length} operation(s).`);
        }
        for (const payload of coalesced) {
          try {
            await this.handleHighlightStudentRowWithRetry(payload);
          } catch (error) {
            // handleHighlightStudentRowWithRetry reports its own errors;
            // never let one payload kill the worker loop.
            console.error("ChromeExtensionService: Unexpected error from highlight handler:", error);
          }
        }
      }
    } finally {
      this._highlightWorkerActive = false;
    }
  }

  /**
   * Resolves a worksheet by name, falling back to date-format variants
   * ("LDA 6-10-2026" ↔ "LDA 06-10-2026", slash/dash separators) when the
   * exact name isn't found. Shared by the single, bulk, and navigate paths.
   * @param {Object} context - Excel.run request context
   * @param {string} targetSheet - Requested sheet name
   * @returns {Promise<Object|null>} Worksheet proxy, or null if not found
   */
  async _resolveWorksheet(context, targetSheet) {
    let worksheet = context.workbook.worksheets.getItemOrNullObject(targetSheet);
    worksheet.load("isNullObject");
    await context.sync();
    if (!worksheet.isNullObject) return worksheet;

    const variations = sheetNameDateVariants(targetSheet);
    if (variations.length <= 1) return null;

    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    for (const variation of variations) {
      const matchingSheet = sheets.items.find(sheet => sheet.name === variation);
      if (matchingSheet) {
        console.log(`ChromeExtensionService: Found sheet "${matchingSheet.name}" using date format normalization (requested: "${targetSheet}")`);
        return context.workbook.worksheets.getItem(matchingSheet.name);
      }
    }
    return null;
  }

  /**
   * Retry wrapper for handleHighlightStudentRow.
   * If the first attempt fails with a retryable error (timeout, stale session),
   * performs up to MAX_HIGHLIGHT_RETRIES retries with exponential backoff.
   * Non-retryable errors (bad params, missing sheet) fail immediately.
   * Executions are serialized by the highlight queue, so the shared retry
   * counter is only ever touched by one invocation at a time.
   * @param {Object} payload - Highlight request payload (same as handleHighlightStudentRow)
   */
  async handleHighlightStudentRowWithRetry(payload) {
    this.highlightRetryCount = 0;

    // Run an Excel health check first if we haven't checked recently (within 2 minutes)
    const healthCheckAge = this.lastExcelHealthCheck ? Date.now() - this.lastExcelHealthCheck : Infinity;
    if (healthCheckAge > 120000) {
      console.log("ChromeExtensionService: Running Excel health check before highlight...");
      const healthy = await this.checkExcelHealth();
      if (!healthy) {
        console.warn("ChromeExtensionService: Excel API health check failed. Will still attempt highlight.");
      }
    }

    while (this.highlightRetryCount <= this.MAX_HIGHLIGHT_RETRIES) {
      try {
        await this.handleHighlightStudentRow(payload);
        return; // Success - exit retry loop
      } catch (error) {
        if (this.isRetryableError(error) && this.highlightRetryCount < this.MAX_HIGHLIGHT_RETRIES) {
          this.highlightRetryCount++;
          const backoffMs = Math.pow(2, this.highlightRetryCount) * 1000; // 2s, 4s, 8s
          console.warn(
            `ChromeExtensionService: Highlight attempt ${this.highlightRetryCount} failed with retryable error: ${error.message}. ` +
            `Retrying in ${backoffMs / 1000}s... (${this.highlightRetryCount}/${this.MAX_HIGHLIGHT_RETRIES})`
          );
          await new Promise(resolve => setTimeout(resolve, backoffMs));
        } else {
          // Non-retryable or out of retries - error already sent to extension by handleHighlightStudentRow
          if (this.highlightRetryCount >= this.MAX_HIGHLIGHT_RETRIES) {
            console.error(
              `ChromeExtensionService: Highlight failed after ${this.MAX_HIGHLIGHT_RETRIES} retries. ` +
              "The Excel API session may need a page refresh."
            );
          }
          return;
        }
      }
    }
  }

  /**
   * Handle highlighting a student row based on extension request
   * Runs in the background - no side panel required.
   * Throws retryable errors so the retry wrapper can catch them.
   * @param {Object} payload - Highlight request payload
   * @param {string} payload.studentName - Student's full name
   * @param {string} payload.syStudentId - Student's SyStudentID
   * @param {number|string} payload.startCol - Starting column (0-based index OR column name)
   * @param {number|string} payload.endCol - Ending column (0-based index OR column name)
   * @param {string} payload.targetSheet - Name of the worksheet to highlight in
   * @param {string} [payload.color='yellow'] - Hex color code for highlight (optional)
   * @param {number|string} [payload.editColumn] - Column to edit (0-based index OR column name, optional)
   * @param {string} [payload.editText] - Text to set in the edit column (optional)
   */
  /**
   * Handle a batched highlight request containing multiple students on the
   * same target sheet. Performs the entire operation in a single Excel.run
   * with one usedRange load and one final context.sync, instead of one
   * round-trip per student.
   */
  async handleBulkHighlightStudentRows(payload) {
    const { students, startCol, endCol, targetSheet, color = '#FFFF00', editColumn } = payload;

    if (!targetSheet) {
      const msg = "Bulk highlight missing targetSheet.";
      console.error("ChromeExtensionService:", msg);
      students.forEach(s => this.sendHighlightConfirmation(s?.syStudentId, "error", msg));
      return;
    }
    if (startCol === undefined || endCol === undefined) {
      const msg = "Bulk highlight missing startCol/endCol.";
      console.error("ChromeExtensionService:", msg);
      students.forEach(s => this.sendHighlightConfirmation(s?.syStudentId, "error", msg));
      return;
    }

    try {
      await this.excelRunWithTimeout(async (context) => {
        const worksheet = await this._resolveWorksheet(context, targetSheet);
        if (!worksheet) {
          const msg = `Sheet "${targetSheet}" not found in the workbook.`;
          console.error("ChromeExtensionService:", msg);
          students.forEach(s => this.sendHighlightConfirmation(s?.syStudentId, "error", msg));
          return;
        }

        const usedRange = worksheet.getUsedRange();
        usedRange.load(["values", "rowCount", "columnCount", "rowIndex"]);
        await context.sync();

        const values = usedRange.values;
        const headers = values[0];

        const resolveColumnIndex = (column) => {
          if (typeof column === 'number') {
            return (column >= 0 && column < headers.length) ? column : -1;
          }
          const name = String(column).trim().toLowerCase();
          for (let col = 0; col < headers.length; col++) {
            if (String(headers[col]).trim().toLowerCase() === name) return col;
          }
          return -1;
        };

        const startColIndex = resolveColumnIndex(startCol);
        const endColIndex = resolveColumnIndex(endCol);
        const editColumnIndex = editColumn !== undefined ? resolveColumnIndex(editColumn) : -1;

        if (startColIndex === -1 || endColIndex === -1 || startColIndex > endColIndex) {
          const msg = `Invalid startCol/endCol in bulk highlight.`;
          console.error("ChromeExtensionService:", msg, { startCol, endCol });
          students.forEach(s => this.sendHighlightConfirmation(s?.syStudentId, "error", msg));
          return;
        }

        // Find ID column once
        const idAliases = ['Student ID', 'SyStudentID', 'Student identifier', 'ID'];
        let idColumnIndex = -1;
        for (let col = 0; col < headers.length; col++) {
          if (idAliases.some(alias => String(headers[col]).trim().toLowerCase() === alias.toLowerCase())) {
            idColumnIndex = col;
            break;
          }
        }
        if (idColumnIndex === -1) {
          const msg = `Could not find a Student ID column in sheet "${targetSheet}".`;
          console.error("ChromeExtensionService:", msg);
          students.forEach(s => this.sendHighlightConfirmation(s?.syStudentId, "error", msg));
          return;
        }

        // Build a single-pass index of syStudentId -> row (avoids N*rows scans)
        const idToRow = new Map();
        for (let row = 1; row < values.length; row++) {
          const cellValue = String(values[row][idColumnIndex]).trim();
          if (cellValue) idToRow.set(cellValue, row);
        }

        const columnCount = endColIndex - startColIndex + 1;
        const successes = [];
        const notFound = [];

        // Suppress StudentView's Outreach change handler for the whole batch
        // so these programmatic writes don't trigger spurious history comments.
        try {
          window.__srkSuppressOutreachHandler = true;
          setTimeout(() => { window.__srkSuppressOutreachHandler = false; }, 2000);
        } catch (e) { /* noop */ }

        for (const student of students) {
          const { studentName, syStudentId, editText } = student;
          const targetRowIndex = idToRow.get(String(syStudentId).trim());
          if (targetRowIndex === undefined) {
            notFound.push(student);
            continue;
          }
          const actualRowIndex = usedRange.rowIndex + targetRowIndex;

          const highlightRange = worksheet.getRangeByIndexes(actualRowIndex, startColIndex, 1, columnCount);
          highlightRange.format.fill.color = color;

          if (editColumnIndex !== -1 && editText !== undefined) {
            const editCell = worksheet.getRangeByIndexes(actualRowIndex, editColumnIndex, 1, 1);
            editCell.values = [[editText]];
          }

          successes.push({ studentName, syStudentId });
        }

        // Single sync commits every highlight and edit in the batch.
        await context.sync();

        console.log(`ChromeExtensionService: Bulk highlight complete — ${successes.length} succeeded, ${notFound.length} not found, in sheet "${targetSheet}".`);

        successes.forEach(s => this.sendHighlightConfirmation(s.syStudentId, "success", "Row highlighted successfully"));
        notFound.forEach(s => this.sendHighlightConfirmation(
          s.syStudentId,
          "error",
          `Student with ID "${s.syStudentId}" not found in sheet "${targetSheet}".`
        ));

        if (successes.length > 0) {
          this.notifyListeners({
            type: "highlight_complete",
            data: {
              bulk: true,
              count: successes.length,
              targetSheet,
              students: successes,
              timestamp: new Date().toISOString()
            }
          });
        }
      });
    } catch (error) {
      console.error("ChromeExtensionService: Error in bulk highlight:", error);
      // Only report errors to the extension when this is the final outcome.
      // The retry wrapper will re-run retryable failures, and a premature
      // error confirmation would mark students failed that later succeed.
      const willRetry = this.isRetryableError(error) && this.highlightRetryCount < this.MAX_HIGHLIGHT_RETRIES;
      if (!willRetry) {
        const msg = error.message || "Unknown error during bulk highlight.";
        students.forEach(s => this.sendHighlightConfirmation(s?.syStudentId, "error", msg));
      }
      // Rethrow so the retry wrapper can catch retryable errors.
      throw error;
    }
  }

  async handleHighlightStudentRow(payload) {
    // Validate Excel is available
    if (typeof window.Excel === "undefined") {
      const msg = "Excel API is not available. The add-in may not be running inside Excel.";
      console.warn("ChromeExtensionService:", msg);
      if (payload && Array.isArray(payload.students)) {
        payload.students.forEach(s => this.sendHighlightConfirmation(s?.syStudentId, "error", msg));
      } else {
        this.sendHighlightConfirmation(payload?.syStudentId, "error", msg);
      }
      return;
    }

    // Bulk highlight path — when the extension sends a batched payload, or
    // the highlight queue coalesced several queued requests into one.
    if (payload && Array.isArray(payload.students) && payload.students.length > 0) {
      return this.handleBulkHighlightStudentRows(payload);
    }

    // Validate required parameters
    if (!payload || !payload.syStudentId || !payload.targetSheet) {
      const msg = "Missing required parameters: syStudentId and targetSheet are required.";
      console.error("ChromeExtensionService:", msg, payload);
      this.sendHighlightConfirmation(payload?.syStudentId, "error", msg);
      return;
    }

    const { studentName, syStudentId, startCol, endCol, targetSheet, color = '#FFFF00', editColumn, editText } = payload;

    // Validate column parameters (can be number or string)
    if (startCol === undefined || startCol === null) {
      const msg = "startCol is required but was not provided.";
      console.error("ChromeExtensionService:", msg, { startCol });
      this.sendHighlightConfirmation(syStudentId, "error", msg);
      return;
    }

    if (endCol === undefined || endCol === null) {
      const msg = "endCol is required but was not provided.";
      console.error("ChromeExtensionService:", msg, { endCol });
      this.sendHighlightConfirmation(syStudentId, "error", msg);
      return;
    }

    // Validate types
    if (typeof startCol !== 'number' && typeof startCol !== 'string') {
      const msg = `startCol must be a number or string, but received ${typeof startCol}.`;
      console.error("ChromeExtensionService:", msg, { startCol });
      this.sendHighlightConfirmation(syStudentId, "error", msg);
      return;
    }

    if (typeof endCol !== 'number' && typeof endCol !== 'string') {
      const msg = `endCol must be a number or string, but received ${typeof endCol}.`;
      console.error("ChromeExtensionService:", msg, { endCol });
      this.sendHighlightConfirmation(syStudentId, "error", msg);
      return;
    }

    if (editColumn !== undefined && typeof editColumn !== 'number' && typeof editColumn !== 'string') {
      const msg = `editColumn must be a number or string, but received ${typeof editColumn}.`;
      console.error("ChromeExtensionService:", msg, { editColumn });
      this.sendHighlightConfirmation(syStudentId, "error", msg);
      return;
    }

    // If editColumn is provided, editText should also be provided
    if (editColumn !== undefined && editText === undefined) {
      console.warn("ChromeExtensionService: editColumn specified but editText is missing");
    }

    try {
      await this.excelRunWithTimeout(async (context) => {
        // Try to get the target worksheet with resilient matching
        const worksheet = await this._resolveWorksheet(context, targetSheet);

        if (!worksheet) {
          const msg = `Sheet "${targetSheet}" not found in the workbook. Verify the sheet name is correct (sheet names are case-sensitive).`;
          console.error("ChromeExtensionService:", msg);
          this.sendHighlightConfirmation(syStudentId, "error", msg);
          return;
        }

        // Load the used range to search for student
        const usedRange = worksheet.getUsedRange();
        usedRange.load(["values", "rowCount", "columnCount", "rowIndex"]);
        await context.sync();

        const values = usedRange.values;
        const headers = values[0]; // First row is headers

        // Helper function to resolve column name to index
        // Returns { index, error } where error is a message string if resolution failed
        const resolveColumnIndex = (column, paramName) => {
          // If it's already a number, validate and return it
          if (typeof column === 'number') {
            if (column < 0) {
              const msg = `${paramName} must be >= 0, but received ${column}.`;
              console.error("ChromeExtensionService:", msg, { [paramName]: column });
              return { index: -1, error: msg };
            }
            if (column >= headers.length) {
              const msg = `${paramName} index (${column}) exceeds the number of columns in the sheet (${headers.length}).`;
              console.error("ChromeExtensionService:", msg, { [paramName]: column });
              return { index: -1, error: msg };
            }
            return { index: column, error: null };
          }

          // It's a string, find the column by name
          const columnName = String(column).trim();
          for (let col = 0; col < headers.length; col++) {
            if (String(headers[col]).trim().toLowerCase() === columnName.toLowerCase()) {
              return { index: col, error: null };
            }
          }

          // Column name not found
          const msg = `Column "${columnName}" not found in sheet "${targetSheet}". Check that the column header exists in row 1.`;
          console.error("ChromeExtensionService:", msg, { [paramName]: column });
          return { index: -1, error: msg };
        };

        // Resolve startCol, endCol, and editColumn to indices
        const startColResult = resolveColumnIndex(startCol, 'startCol');
        const endColResult = resolveColumnIndex(endCol, 'endCol');

        if (startColResult.error || endColResult.error) {
          const msg = startColResult.error || endColResult.error;
          this.sendHighlightConfirmation(syStudentId, "error", msg);
          return;
        }

        const startColIndex = startColResult.index;
        const endColIndex = endColResult.index;

        if (startColIndex > endColIndex) {
          const msg = `startCol (${startCol}, index ${startColIndex}) must be less than or equal to endCol (${endCol}, index ${endColIndex}).`;
          console.error("ChromeExtensionService:", msg);
          this.sendHighlightConfirmation(syStudentId, "error", msg);
          return;
        }

        let editColumnIndex = -1;
        if (editColumn !== undefined) {
          const editColResult = resolveColumnIndex(editColumn, 'editColumn');
          if (editColResult.error) {
            this.sendHighlightConfirmation(syStudentId, "error", editColResult.error);
            return;
          }
          editColumnIndex = editColResult.index;
        }

        // Find the ID column index
        let idColumnIndex = -1;
        const idAliases = ['Student ID', 'SyStudentID', 'Student identifier', 'ID'];

        for (let col = 0; col < headers.length; col++) {
          if (idAliases.some(alias =>
            String(headers[col]).trim().toLowerCase() === alias.toLowerCase()
          )) {
            idColumnIndex = col;
            break;
          }
        }

        if (idColumnIndex === -1) {
          const msg = `Could not find a Student ID column in sheet "${targetSheet}". Expected one of: ${idAliases.join(', ')}.`;
          console.error("ChromeExtensionService:", msg);
          this.sendHighlightConfirmation(syStudentId, "error", msg);
          return;
        }

        // Find the row with matching syStudentId
        let targetRowIndex = -1;
        for (let row = 1; row < values.length; row++) { // Start from 1 to skip headers
          const cellValue = String(values[row][idColumnIndex]).trim();
          if (cellValue === String(syStudentId).trim()) {
            targetRowIndex = row;
            break;
          }
        }

        if (targetRowIndex === -1) {
          const msg = `Student with ID "${syStudentId}" not found in sheet "${targetSheet}". Verify the student ID exists in the sheet.`;
          console.warn("ChromeExtensionService:", msg);
          this.sendHighlightConfirmation(syStudentId, "error", msg);
          return;
        }

        // Calculate the actual row index in the worksheet (accounting for used range offset)
        const actualRowIndex = usedRange.rowIndex + targetRowIndex;

        // Calculate column count to highlight
        const columnCount = endColIndex - startColIndex + 1;

        // Get the range to highlight
        const highlightRange = worksheet.getRangeByIndexes(
          actualRowIndex,
          startColIndex,
          1, // Just one row
          columnCount
        );

        // Apply the highlight color
        highlightRange.format.fill.color = color;
        await context.sync();

        // Edit cell if editColumn and editText are provided
        if (editColumn !== undefined && editText !== undefined && editColumnIndex !== -1) {
          const editCell = worksheet.getRangeByIndexes(
            actualRowIndex,
            editColumnIndex,
            1, // One row
            1  // One column
          );
          // Suppress StudentView's Outreach change handler so this programmatic
          // write doesn't trigger an unintended history comment when the student
          // view modal is open during a submission highlight.
          try {
            window.__srkSuppressOutreachHandler = true;
            setTimeout(() => { window.__srkSuppressOutreachHandler = false; }, 2000);
          } catch (e) { /* noop */ }
          editCell.values = [[editText]];
          await context.sync();

          console.log(`ChromeExtensionService: Successfully edited cell at column "${editColumn}" (index ${editColumnIndex}) with text "${editText}"`);
        }

        console.log(`ChromeExtensionService: Successfully highlighted student "${studentName}" (ID: ${syStudentId}) in "${targetSheet}" from column "${startCol}" (index ${startColIndex}) to "${endCol}" (index ${endColIndex})`);

        // Send confirmation back to Chrome extension
        this.sendHighlightConfirmation(syStudentId, "success", "Row highlighted successfully");

        // Notify listeners of successful highlight
        this.notifyListeners({
          type: "highlight_complete",
          data: {
            studentName,
            syStudentId,
            targetSheet,
            startCol,
            endCol,
            startColIndex,
            endColIndex,
            color,
            editColumn,
            editColumnIndex,
            editText,
            timestamp: new Date().toISOString()
          }
        });
      });
    } catch (error) {
      console.error("ChromeExtensionService: Error highlighting student row:", error);
      console.error("ChromeExtensionService: Error details - code:", error.code, "message:", error.message);

      // Determine a helpful error message based on common failure scenarios
      let errorMessage = error.message || "An unknown error occurred while highlighting the row.";
      let isRetryable = this.isRetryableError(error);

      if (error.code === "Timeout") {
        errorMessage = `Excel API did not respond within ${this.EXCEL_RUN_TIMEOUT_MS / 1000}s. ` +
          "The session may have expired. Try refreshing the page.";
        this.excelHealthy = false;
      } else if (error.code === "InvalidObjectPath") {
        errorMessage = "Excel session appears stale. The API context may have expired. Try refreshing the page.";
        this.excelHealthy = false;
      } else if (error.code === "InvalidReference") {
        errorMessage = "Invalid cell reference. The worksheet structure may have changed.";
      } else if (error.code === "GeneralException") {
        errorMessage = "Excel encountered an error. The workbook may be protected, read-only, or in use by another process.";
      } else if (error.code === "AccessDenied") {
        errorMessage = "Access denied. The workbook or sheet may be protected.";
      } else if (error.message && error.message.toLowerCase().includes("network")) {
        errorMessage = "A network error occurred. Check your connection to Excel Online.";
      }

      // If retryable, re-throw so the retry wrapper can handle it
      if (isRetryable && this.highlightRetryCount < this.MAX_HIGHLIGHT_RETRIES) {
        console.warn(`ChromeExtensionService: Error is retryable (code: ${error.code}). Will retry.`);
        throw error;
      }

      // Non-retryable or final retry exhausted - report to extension
      const retryInfo = this.highlightRetryCount > 0
        ? ` (failed after ${this.highlightRetryCount} retries)`
        : "";
      const finalMessage = errorMessage + retryInfo;

      // Send error confirmation back to Chrome extension
      this.sendHighlightConfirmation(syStudentId, "error", finalMessage);

      // Notify listeners of error
      this.notifyListeners({
        type: "highlight_error",
        data: {
          studentName,
          syStudentId,
          editColumn,
          editText,
          error: finalMessage,
          retryCount: this.highlightRetryCount,
          errorCode: error.code || "unknown",
          timestamp: new Date().toISOString()
        }
      });
    }
  }

  /**
   * Handle navigating to a student's row based on extension request.
   * Activates the target sheet and selects the student's row so that the
   * existing onSelectionChanged listener in StudentView picks it up automatically.
   * No confirmation message is sent back.
   * @param {Object} payload - Navigation request payload
   * @param {string} payload.syStudentId - Student's SyStudentID
   * @param {string} payload.targetSheet - Name of the worksheet to navigate to
   */
  async handleNavigateToStudent(payload) {
    if (typeof window.Excel === "undefined") {
      console.warn("ChromeExtensionService: Excel API is not available for navigation.");
      return;
    }

    if (!payload || !payload.syStudentId || !payload.targetSheet) {
      console.error("ChromeExtensionService: Navigate requires syStudentId and targetSheet.", payload);
      return;
    }

    const { syStudentId, targetSheet } = payload;

    try {
      await Excel.run(async (context) => {
        // --- Resolve the target worksheet (with date-format normalization) ---
        const worksheet = await this._resolveWorksheet(context, targetSheet);

        if (!worksheet) {
          console.error(`ChromeExtensionService: Navigate - sheet "${targetSheet}" not found.`);
          return;
        }

        // --- Find the student row ---
        const usedRange = worksheet.getUsedRange();
        usedRange.load(["values", "rowCount", "columnCount", "rowIndex"]);
        await context.sync();

        const values = usedRange.values;
        const headers = values[0];

        const idAliases = ['Student ID', 'SyStudentID', 'Student identifier', 'ID'];
        let idColumnIndex = -1;
        for (let col = 0; col < headers.length; col++) {
          if (idAliases.some(alias => String(headers[col]).trim().toLowerCase() === alias.toLowerCase())) {
            idColumnIndex = col;
            break;
          }
        }

        if (idColumnIndex === -1) {
          console.error(`ChromeExtensionService: Navigate - no Student ID column found in "${targetSheet}".`);
          return;
        }

        let targetRowIndex = -1;
        for (let row = 1; row < values.length; row++) {
          if (String(values[row][idColumnIndex]).trim() === String(syStudentId).trim()) {
            targetRowIndex = row;
            break;
          }
        }

        if (targetRowIndex === -1) {
          console.warn(`ChromeExtensionService: Navigate - student "${syStudentId}" not found in "${targetSheet}".`);
          return;
        }

        // --- Activate sheet and select the student's row ---
        const actualRowIndex = usedRange.rowIndex + targetRowIndex;
        const studentRow = worksheet.getRangeByIndexes(actualRowIndex, 0, 1, headers.length);

        worksheet.activate();
        studentRow.select();
        await context.sync();

        console.log(`ChromeExtensionService: Navigated to student "${syStudentId}" in "${targetSheet}" (row ${actualRowIndex + 1}).`);
      });
    } catch (error) {
      console.error("ChromeExtensionService: Error navigating to student row:", error);
    }
  }

  /**
   * Send a ping to check if extension is installed
   * Sends to both current window and parent window
   */
  sendPing() {
    const message = { type: "SRK_CHECK_EXTENSION" };

    // Send to self
    window.postMessage(message, "*");

    // Send to parent if in iframe
    if (window.parent && window.parent !== window) {
      window.parent.postMessage(message, "*");
    }
  }

  /**
   * Start pinging to detect the extension
   * @param {number} interval - Milliseconds between pings (default: 2000)
   */
  startPinging(interval = 2000) {
    if (this.isChecking) return;

    console.log("ChromeExtensionService: Starting extension detection...");
    this.isChecking = true;

    // Send initial ping
    this.sendPing();

    // Setup interval for periodic pings
    this.pingInterval = setInterval(() => {
      if (!this.isExtensionInstalled) {
        this.sendPing();
      } else {
        this.stopPinging();
      }
    }, interval);
  }

  /**
   * Stop pinging for extension detection
   */
  stopPinging() {
    if (this.pingInterval) {
      console.log("ChromeExtensionService: Stopping extension detection.");
      clearInterval(this.pingInterval);
      this.pingInterval = null;
      this.isChecking = false;
    }
  }

  /**
   * Start keep-alive heartbeat for Chrome extension context
   * This prevents the extension from going dormant.
   * Also starts the message listener watchdog to detect stale listeners.
   * @param {number} interval - Milliseconds between pings (default: 20000)
   */
  startKeepAlive(interval = 20000) {
    if (this.keepAliveInterval) return;

    console.log("ChromeExtensionService: Starting keep-alive heartbeat...");

    this.keepAliveInterval = setInterval(() => {
      if (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.sendMessage) {
        chrome.runtime.sendMessage({ type: 'keep_alive_ping' }, (response) => {
          if (chrome.runtime.lastError) {
            // Ignore errors - extension might not be installed
          }
        });
      }
    }, interval);

    // Also start the message listener watchdog
    this.startMessageListenerWatchdog();
  }

  /**
   * Stop keep-alive heartbeat
   */
  stopKeepAlive() {
    if (this.keepAliveInterval) {
      console.log("ChromeExtensionService: Stopping keep-alive heartbeat.");
      clearInterval(this.keepAliveInterval);
      this.keepAliveInterval = null;
    }
  }

  /**
   * Add a listener for extension events
   * @param {Function} callback - Function to call when events occur
   * @returns {Function} Cleanup function to remove the listener
   */
  addListener(callback) {
    this.listeners.add(callback);

    // Return cleanup function
    return () => {
      this.listeners.delete(callback);
    };
  }

  /**
   * Notify all listeners of an event
   * @param {Object} event - Event object to send to listeners
   */
  notifyListeners(event) {
    this.listeners.forEach(listener => {
      try {
        listener(event);
      } catch (error) {
        console.error("ChromeExtensionService: Listener error:", error);
      }
    });
  }

  /**
   * Send a custom message to the extension
   * @param {Object} message - Message object to send
   */
  sendMessage(message) {
    // If there's a parent window (Office Add-in running in iframe), send only to parent
    // This prevents duplicate messages to the Chrome extension
    if (window.parent && window.parent !== window) {
      console.log(`📤 ChromeExtensionService: Sending message (${message.type}) to parent window:`, message);
      window.parent.postMessage(message, "*");
    } else {
      // No parent window, send to self
      console.log(`📤 ChromeExtensionService: Sending message (${message.type}) to window:`, message);
      window.postMessage(message, "*");
    }
  }

  /**
   * Send a message using chrome.runtime API (if available)
   * @param {Object} message - Message to send
   * @param {Function} callback - Optional callback for response
   */
  sendChromeMessage(message, callback) {
    if (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.sendMessage) {
      chrome.runtime.sendMessage(message, (response) => {
        if (chrome.runtime.lastError) {
          console.warn("ChromeExtensionService: Chrome message error:", chrome.runtime.lastError);
        }
        if (callback) callback(response);
      });
    } else {
      console.warn("ChromeExtensionService: chrome.runtime.sendMessage not available");
    }
  }

  /**
   * Send selected student data to the Chrome extension
   * @param {Array|Object} students - Single student object or array of students
   * @param {string|null} directPhone - Phone number from a directly-selected cell (fast-path)
   * @param {boolean} autoCall - Whether the extension should automatically initiate the call
   */
  sendSelectedStudents(students, directPhone = null, autoCall = false) {
    // Normalize to array format
    const studentArray = Array.isArray(students) ? students : [students];

    // Extract only the required fields
    const payload = studentArray.map(student => ({
      name: student.StudentName || student.name || "",
      syStudentId: student.ID || student.syStudentId || "",
      phone: student.Phone || student.phone || "",
      otherPhone: student.OtherPhone || student.otherPhone || "",
      ...(student.directPhone ? { directPhone: student.directPhone } : {}),
      ...(student.isOtherContact ? { isOtherContact: true } : {})
    }));

    const message = {
      type: "SRK_SELECTED_STUDENTS",
      data: {
        students: payload,
        count: payload.length,
        timestamp: new Date().toISOString(),
        directPhone: directPhone || null,  // Phone number from single cell selection
        autoCall: autoCall  // Auto-initiate call (used by ribbon call button)
      }
    };

    console.log("ChromeExtensionService: Sending selected students:", message);
    this.sendMessage(message);
  }

  /**
   * Get current extension installation status
   * @returns {boolean} Whether the extension is detected as installed
   */
  getInstallationStatus() {
    return this.isExtensionInstalled;
  }

  /**
   * Reset the service state (useful for testing or reinitializing)
   */
  reset() {
    this.stopPinging();
    this.stopKeepAlive();
    this.stopMessageListenerWatchdog();
    this.isExtensionInstalled = false;
    this.isChecking = false;
    this.excelHealthy = null;
    this.lastExcelHealthCheck = null;
    this.highlightRetryCount = 0;
    // Drop any queued highlight requests (an in-flight one finishes on its own)
    this._highlightQueue.length = 0;
  }

  /**
   * Cleanup all intervals and listeners
   */
  cleanup() {
    this.stopPinging();
    this.stopKeepAlive();
    this.stopMessageListenerWatchdog();
    window.removeEventListener("message", this.handleMessage);
    this._messageListenerAttached = false;
    this.listeners.clear();
    this._highlightQueue.length = 0;
  }
}

// Create and export a singleton instance
const chromeExtensionService = new ChromeExtensionService();

export default chromeExtensionService;
