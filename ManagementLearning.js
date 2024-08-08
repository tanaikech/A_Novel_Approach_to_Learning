/**
 * Class object for management learning.
 * @class
 */
class ManagementLearning {

  constructor() {
    /** @private */
    this.startTime = new Date();

    /**
     * Function name of client function. Default name is "main_ManagementLearning".
     * @private
    */
    this.mainFunctionName = "main_ManagementLearning";

    /**
     * Default values of dashboard.
     * @private
    */
    this.object = {
      apiKey: null,
      useAccessToken: false,
      model: "models/gemini-1.5-flash-latest",
      version: "v1beta",
      field: "",
      targetAccuracyRate: 80,
      accessToken: null,
      questionsInGoogleForm: 10,
      repetitionsForQuestion: 1,
      repeatCount: 5,
      useGemini: true,
      googleFormPublishedUrl: "",
      googleFormEditUrl: "",
      categories: "",
      currentAccuracyRate: 0,
      status: `Init state. Please set "field" and run the script with 'Run' from a custom menu. The script generates questions.`,
    };

    /** @private */
    this.accessToken = ScriptApp.getOAuthToken();

    /** @private */
    this.spreadsheet = null;

    /** @private */
    this.dashboardSheetName = "dashboard";

    /** @private */
    this.dashboardSheetRowHead = [];

    /** @private */
    this.commentsSheetName = "comments";

    /** @private */
    this.questionsSheetName = "questions";

    /** @private */
    this.questionsByUserSheetName = "questionsByUser";

    /** @private */
    this.archiveSheetName = "archive";

    /** @private */
    this.dashboardSheet = null;

    /** @private */
    this.commentsSheet = null;

    /** @private */
    this.questionsSheet = null;

    /** @private */
    this.questionsByUserSheet = null;

    /** @private */
    this.archiveSheet = null;

    /** @private */
    this.headersQuestionsSheet = ["cycle", "category", "question", "choices", "answer", "countCorrects", "countIncorrects", "accuracyRate", "history"];

    /** @private */
    this.delimiter1 = "\n";

    /** @private */
    this.objToGemini = null;
  }


  /**
   * ### Description
   * Main method.
   *
   * @param {Object} object Event object of installable OnSubmit trigger on Google Forms.
   * @param {SpreadsheetApp.Range} object.range
   * @param {FormApp.Form} object.source
   * @param {FormApp.FormResponse} object.response
   * @param {String} object.triggerUid
   * @param {String} object.authMode
   * 
   * @return {void}
   *
   */
  run(object = null) {
    console.log(`Get user's inputted values from Spreadsheet.`);
    this.getSheets_();
    this.getInitParams_();

    // Retrieve questions.
    const srcRange = this.questionsSheet.getDataRange();
    const [, ...values] = srcRange.getValues();
    const cycle = values.length > 0 ? values[values.length - 1][0] : 1;

    if (object && ['authMode', 'response', 'source', 'triggerUid'].every(k => object[k])) {
      this.processGoogleForm_({ object, srcRange, values });
    } else {
      this.setOnSubmitTrigger_();
    }
    this.object.form.setCustomClosedFormMessage("Now updating questions in Google Form. Please wait a minute.").setAcceptingResponses(false);
    let selectedQuestions = [];
    if (values.length == 0) {
      console.log(`In the initial situation, generate questions using Gemini and randomly select questions.`);
      selectedQuestions = this.initSituation_(cycle);
    } else {
      const parsedValues = this.parseValues_(values);
      const sumAccuracyRate = parsedValues.reduce((t, { averageAccuracyRate }) => (t += averageAccuracyRate, t), 0);
      const numCategories = parsedValues.length;
      this.object.currentAccuracyRate = sumAccuracyRate / numCategories;
      this.setValuesToDashboardSheet_();
      if (this.object.targetAccuracyRate < this.object.currentAccuracyRate) {
        if (this.object.repeatCount > cycle) {
          console.log(`Under the value "currentAccuracyRate" is over the value "targetAccuracyRate" and the value "cycle" is less than the value "repeatCount", generate questions using Gemini and select questions.`);

          console.log(`Generate summary for the answers.`);
          const [, ...values] = srcRange.getValues();
          const reqiredCols = ["category", "question", "history"].map(h => this.headersQuestionsSheet.indexOf(h));
          const questionsCSV = [this.headersQuestionsSheet, ...values.filter(([a]) => a == cycle)]
            .map(r => reqiredCols.map(c => r[c]))
            .map(r => r.map(c => isNaN(c) ? `"${c.replace(/"/g, '\\"')}"` : c).join(",")).join("\n");
          const summary = this.summaryAnswersByGemini_({ questionsCSV });
          this.commentsSheet.appendRow([this.startTime, summary.result || "No result was returned.", cycle])
          selectedQuestions = this.overTargetAccuracyRate_(cycle + 1);
        } else {
          console.log(`The value "currentAccuracyRate" is over the value "targetAccuracyRate" and the value "cycle" is reached to the value "repeatCount". By this, your learning was finished. Google Form is closed.`);
          this.object.form.setCustomClosedFormMessage("Congratulations! Your learning was finished with your goal.").setAcceptingResponses(false);

          console.log(`Generate summary for all answers.`);
          const [, ...values] = srcRange.getValues();
          const reqiredCols = ["category", "question", "history"].map(h => this.headersQuestionsSheet.indexOf(h));
          const questionsCSV = [this.headersQuestionsSheet, ...values]
            .map(r => reqiredCols.map(c => r[c]))
            .map(r => r.map(c => isNaN(c) ? `"${c.replace(/"/g, '\\"')}"` : c).join(",")).join("\n");
          const summary = this.summaryAnswersByGemini_({ questionsCSV });
          this.commentsSheet.appendRow([this.startTime, summary.result || "No result was returned.", "Finished!"]);
          this.object.status = `Congratulations! Your learning was finished. If you want to do other learning, please "Reset" and "Run" again.`;
          this.setValuesToDashboardSheet_();
          this.deleteCurrentOnSubmitTrigger_();
          return;
        }
      } else {
        console.log(`Under the value "currentAccuracyRate" is less than the value "targetAccuracyRate", select questions.`);
        selectedQuestions = this.underTargetAccuracyRate_({ values, cycle });
      }
    }

    // Set questions in Google Form.
    this.setQuestionsToGooglForm_(selectedQuestions);
    this.object.form.setCustomClosedFormMessage("").setAcceptingResponses(true);
    this.object.status = `Questions were prepared. You can learn about "${this.object.field}". Please open the Google Form and answer them.`;
    this.setValuesToDashboardSheet_();

    console.log("Done.");
  }

  /**
   * ### Description
   * Reset.
   *
   * @return {void}
   */
  reset() {
    console.log("Start resetting.");
    this.getSheets_();

    // Reset question sheet.
    const srcRange1 = this.questionsSheet.getRange(2, 1, this.questionsSheet.getLastRow() - 1, this.questionsSheet.getLastColumn());
    const values1 = srcRange1.getValues();
    // const values1b = [["Reset at", new Date(), ...Array(values1[0].length - 2).fill(null)], ...values1];
    this.archiveSheet.getRange(this.archiveSheet.getLastRow() + 1, 1, values1.length, values1[0].length).setValues(values1);
    srcRange1.clearContent();

    // Reset dashboard sheet.
    const srcRange2 = this.dashboardSheet.getRange(2, 1, this.dashboardSheet.getLastRow() - 1, this.dashboardSheet.getLastColumn());
    const values2 = srcRange2.getValues();
    const rh = new Set(["categories", "googleFormPublishedUrl", "googleFormEditUrl", "currentAccuracyRate", "status"]);
    const values2b = values2.map(r => {
      const [a, , ...cd] = r;
      return rh.has(a) ? [a, a == "status" ? `Init state. Please set "field" and run the script with 'Run' from a custom menu. The script generates questions.` : null, ...cd] : r;
    });
    srcRange2.setValues(values2b);
    console.log("Done.");
  }

  /**
   * ### Description
   * Get work sheets.
   * 
   * @return {void}
   *
   * @private
   */
  getSheets_() {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    [
      this.dashboardSheet,
      this.commentsSheet,
      this.questionsSheet,
      this.questionsByUserSheet,
      this.archiveSheet,
    ] = [
      this.dashboardSheetName,
      this.commentsSheetName,
      this.questionsSheetName,
      this.questionsByUserSheetName,
      this.archiveSheetName,
    ]
      .map(s => this.spreadsheet.getSheetByName(s) || this.spreadsheet.insertSheet(s));
  }

  /**
   * ### Description
   * Get user's values.
   * 
   * @return {void}
   *
   * @private
   */
  getInitParams_() {
    const range = this.dashboardSheet.getDataRange();
    const v = range.getValues();
    const [head, ...values] = v;
    this.dashboardSheetRowHead = values.map(([a]) => a);
    const keys = Object.keys(this.object);

    values.forEach(([a, b]) => {
      const ta = a.trim();
      const tb = typeof b == "string" ? b.trim() : b;
      if (keys.includes(ta)) {
        this.object[ta] = tb ?? this.object[ta];
      }
    });
    if (!this.object.field) {
      this.showError_(`"field" is not set. Please set the value of "field" that you want to learn.`);
    }
    this.objToGemini = this.createObjectForGeminiWithFiles_();

    keys.forEach(ta => {
      const tb = this.object[ta];
      if (ta == "categories" && !tb) {
        this.object[ta] = this.generateCategories_().join(",");
      } else if (ta == "targetAccuracyRate") {
        this.object[ta] = (tb == "" || !tb || isNaN(tb)) ? this.object[ta] : tb;
      } else if (ta == "currentAccuracyRate") {
        this.object[ta] = (tb == "" || !tb || isNaN(tb)) ? 0 : tb;
      } else if (ta == "googleFormPublishedUrl" || ta == "googleFormEditUrl") {
        if (tb == "" || !tb) {
          if (!this.object.form) {
            const ssId = this.spreadsheet.getId();
            this.object.form = FormApp
              .create("GoogleForm_for_ManagementLearning")
              .setIsQuiz(true)
              .setShuffleQuestions(true);
            DriveApp.getFileById(this.object.form.getId()).moveTo(DriveApp.getFileById(ssId).getParents().next());
          }
          if (!this.object.googleFormPublishedUrl) {
            this.object.googleFormPublishedUrl = this.object.form.getPublishedUrl();
          }
          if (!this.object.googleFormEditUrl) {
            this.object.googleFormEditUrl = this.object.form.getEditUrl();
          }
        } else if (ta == "googleFormPublishedUrl") {
          this.object.googleFormPublishedUrl = tb;
        } else if (ta == "googleFormEditUrl") {
          this.object.googleFormEditUrl = tb;
          this.object.form = FormApp.openByUrl(tb);
        }
      }
    });

    this.object.form
      .setTitle(`Lets's learn ${this.object.field}!`)
      .setDescription(`The field of learning is ${this.object.field}.`);
    if (this.object.useAccessToken === true) {
      this.object.accessToken = this.accessToken;
    }
    const ar = values.map(([a, _, ...cd]) => [a, this.object[a], ...cd]);
    range.setValues([head, ...ar]);
    SpreadsheetApp.flush();
  }

  /**
   * ### Description
   * Process the event object from Google Form.
   * 
   * @param {Object} obj Object including the event object and the source range.
   * @return {void}
   *
   * @private
   */
  processGoogleForm_(obj) {
    const { object, srcRange, values } = obj;
    const tempObj = object.response.getGradableItemResponses().reduce((o, r) => {
      const item = r.getItem();
      const decodeCode = JSON.parse(Utilities.newBlob(Utilities.base64Decode(JSON.parse(item.getHelpText()).identityCode)).getDataAsString());
      const cycle = decodeCode.cycle;
      o[`${item.getTitle().trim()}_${cycle}`] = { score: r.getScore() };
      return o;
    }, {});
    values.forEach(r => {
      const key = `${r[2].trim()}_${r[0]}`;
      if (tempObj[key]) {
        const { score } = tempObj[key];
        if (score == 1) {
          r[5] += 1;
        } else {
          r[6] += 1;
        }
        r[8] = [...r[8].split(",").map(e => e.trim()).filter(String), score == 1 ? "Correct" : "Incorrect"].join(",");
      }
      if ((r[5] + r[6]) >= this.object.repetitionsForQuestion) {
        r[7] = 100 * r[5] / (r[5] + r[6]);
      } else {
        r[7] = 0;
      }
    });
    srcRange.setValues([this.headersQuestionsSheet, ...values]);
    SpreadsheetApp.flush();
  }

  /**
   * ### Description
   * Install OnSubmit trigger to Google Form.
   * 
   * @return {void}
   *
   * @private
   */
  setOnSubmitTrigger_() {
    this.deleteCurrentOnSubmitTrigger_();
    ScriptApp.newTrigger(this.mainFunctionName).forForm(this.object.form).onFormSubmit().create();
  }

  /**
   * ### Description
   * Delete current OnSubmit trigger.
   * 
   * @return {void}
   *
   * @private
   */
  deleteCurrentOnSubmitTrigger_() {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() == this.mainFunctionName) {
        ScriptApp.deleteTrigger(t);
      }
    });
  }

  /**
   * ### Description
   * Set the selected questions to Googl Form.
   * 
   * @param {Array} selectedQuestions Array including object of the selected questions.
   * @return {void}
   *
   * @private
   */
  setQuestionsToGooglForm_(selectedQuestions) {
    console.log(`Updating Google Form with new questions.`);

    // Reset form
    this.clearForm_();

    const form = this.object.form;
    selectedQuestions.forEach(({ cycle, category, question, choices, answer }) => {
      const base64 = Utilities.base64Encode(JSON.stringify({ cycle, category, question }));
      const text = JSON.stringify({ identityCode: base64 });
      const item = form.addMultipleChoiceItem().setTitle(question).setHelpText(text).setPoints(1);
      const c = choices.split("\n").map(ci => item.createChoice(ci, ci == answer));
      item.setChoices(this.getShuffledArr_(c));
    });
  }

  /**
   * ### Description
   * Set values to dashboard sheet.
   *
   * @private
   */
  setValuesToDashboardSheet_() {
    const values = this.dashboardSheetRowHead.map(a => [this.object[a] || ""]);
    this.dashboardSheet.getRange(2, 2, values.length).setValues(values);
  }

  /**
   * ### Description
   * At initial situation, generate new questions, put them into sheet, and select questions for Google Form.
   * 
   * @param {Number} cycle Number of cycle for "repeatCount".
   * @return {Array} Array including questions.
   *
   * @private
   */
  initSituation_(cycle) {
    const generatedQuestions = this.generateQuestions_main_();
    const array = this.convertGeminiObjectToArray_({ generatedQuestions, cycle });
    const dstValues = [this.headersQuestionsSheet, ...array];
    this.questionsSheet.getRange(1, 1, dstValues.length, dstValues[0].length).setValues(dstValues);
    const selectedQuestions = this.getShuffledArr_(array).splice(0, this.object.questionsInGoogleForm);
    return this.convertArrayToGeminiObject_(selectedQuestions);
  }

  /**
   * ### Description
   * Select questions under the condition that the value "currentAccuracyRate" is less than the goal.
   * 
   * @param {Object} object Object including all questions at the latest "cycle".
   * @return {Array} Array including questions.
   *
   * @private
   */
  underTargetAccuracyRate_(object) {
    const { values, cycle } = object;
    return this.object.useGemini
      ? this.selectQuestionsByGemini_({ values, cycle }) // Select questions by Gemini
      : this.selectQuestionsByRandom_({ values, cycle }); // Select questions by random
  }

  /**
   * ### Description
   * Generate new questions and put them into sheet.
   * Select questions under the condition that the value "currentAccuracyRate" is over the goal.
   * 
   * @param {Number} cycle Number of cycle for "repeatCount".
   * @return {Array} Array including questions.
   *
   * @private
   */
  overTargetAccuracyRate_(cycle) {
    const generatedQuestions = this.generateQuestions_main_();
    const dstValues = this.convertGeminiObjectToArray_({ generatedQuestions, cycle });
    this.questionsSheet.getRange(this.questionsSheet.getLastRow() + 1, 1, dstValues.length, dstValues[0].length).setValues(dstValues);
    const selectedQuestions = this.getShuffledArr_(dstValues).splice(0, this.object.questionsInGoogleForm);
    return this.convertArrayToGeminiObject_(selectedQuestions);
  }

  /**
   * ### Description
   * Convert an object from Gemini to an array.
   * 
   * @param {Object} object Object including the generated questions and the value of "cycle".
   * @return {Array} Array including questions.
   *
   * @private
   */
  convertGeminiObjectToArray_(object) {
    const { generatedQuestions, cycle } = object;
    return generatedQuestions.reduce((ar, { category, questions }) => {
      questions.forEach(({ question, choices, answer }) => {
        ar.push([cycle, category, question, choices.join(this.delimiter1), choices[answer], 0, 0, 0, ""]);
      });
      return ar;
    }, []);
  }

  /**
   * ### Description
   * Convert an array to an object for Gemini.
   * 
   * @param {Array} array Array including questions.
   * @return {Array} Array including questions.
   *
   * @private
   */
  convertArrayToGeminiObject_(array) {
    return array.map(r => this.headersQuestionsSheet.reduce((o, c, j) => (o[c] = r[j], o), {}));
  }

  /**
   * ### Description
   * Parse 2 dimensional array of Spreadsheet to an object.
   * 
   * @param {Array} values Array including questions including answers.
   * @return {Array} Array including questions.
   * 
   * @private
   */
  parseValues_(values) {
    return Array.from(values.reduce((m, [cycle, category, question, choices, answer, countCorrects, countIncorrects, accuracyRate, history]) => {
      const cat = category.trim();
      const c = choices.split(this.delimiter1).map(e => e.trim());
      const temp = { cycle, category: cat, question, choices: c, answer: c.indexOf(answer), countCorrects, countIncorrects, accuracyRate, history };
      return m.set(cat, m.has(cat) ? [...m.get(cat), temp] : [temp]);
    }, new Map()))
      .map(([category, questions]) => {
        const averageAccuracyRate = questions.reduce((t, o) => (t += o.accuracyRate, t), 0) / questions.length;
        return { category, questions, averageAccuracyRate };
      });
  }


  /**
   * ### Description
   * Generate categories.
   * 
   * @return {Array} Generated categories.
   *
   * @private
   */
  generateCategories_() {
    console.log(`Generate categories from your inputted field "${this.object.field}" by Gemini.`);
    try {
      const jsonSchema = {
        title: "Create category",
        description: `I would like to learn ${this.object.field}. In order to learn ${this.object.field}, what are the categories related to ${this.object.field}? At least, propose suitable 10 categories in order of importance category. Propose only 10 or fewer suitable categories.`,
        type: "array",
        items: {
          type: "string",
          description: `Created category.`,
        },
      };
      const g = new GeminiWithFiles(this.objToGemini);
      let res = g.generateContent({ jsonSchema });
      if (res.length > 10) {
        res = res.splice(0, 10);
      }
      console.log(`Wait for 5 seconds to use Gemini API.`);
      Utilities.sleep(5000);
      return res;
    } catch ({ stack }) {
      this.showError_(stack);
    }
  }

  /**
   * ### Description
   * Generate questions from categories.
   * 
   * @return {Object} Generated questions.
   *
   * @private
   */
  generateQuestions_main_(retry = 3) {
    console.log(`Generate questions about your inputted field "${this.object.field}" by Gemini.`);
    const userData = this.questionsByUserSheet.getDataRange().getValues();
    let userDataBlob = null;
    if (userData.length > 1) {
      const uderDataCSV = userData.map(r => r.map(c => isNaN(c) ? `"${c.replace(/"/g, '\\"')}"` : c).join(",")).join("\n");
      userDataBlob = Utilities.newBlob(uderDataCSV, MimeType.PLAIN_TEXT, "userData.txt");
    }

    const [, ...values] = this.questionsSheet.getDataRange().getValues();
    const ar = values.map(r => {
      const choices = r[1].trim().split(",").map(s => s.trim());
      return { question: r[0], choices, answer: choices.indexOf(r[2].trim()) };
    });
    const blob = ar.length > 0 ? Utilities.newBlob(JSON.stringify(ar), MimeType.PLAIN_TEXT, "questionsByGemini.txt") : null;

    // Here, in the current stage, the number of questions for each category is the same with the value of "this.object.repetitionsForQuestion".
    const numberOfCreateQuestions = this.object.repetitionsForQuestion;

    const jsonSchema = {
      title: "Create questions and answers.",
      description: [
        `I want to learn and understand "${this.object.field}" deeply. In order to achieve this, I created the following categories for ${this.object.field}.`,
        `<Categories>${this.object.categories}</Categories>`,
        `Using these categories, please create ${numberOfCreateQuestions} questions for every category to help me learn and understand ${this.object.field} deeply.`,
        `Each question is required to have 5 choices including the correct answer. Of course, include the correct answer.`,
        `Difficult questions for helping understand ${this.object.field} are welcome.`,
        blob ? `Although you can use these questions of the JSON including the questions and answers in the following file (the filename is questionsByGemini.txt.), create questions by updating without directly duplicating the JSON.` : "",
        `When you use the methods for teaching about script, please don't use nonexistent methods.`,
        `The final goal is to get an ${this.object.targetAccuracyRate} % accuracy rate.`,
        userDataBlob ? `Questions in the following file (the filename is userData.txt.) are required to be included in the generated questions. But, please proofread those questions and include the generated questions.` : "",
      ].join("\n"),
      type: "array",
      items: {
        type: "object",
        properties: {
          category: { description: "Category", type: "string" },
          questions: {
            description: "Questions for the category",
            type: "array",
            items: {
              description: "Object including generated question, choices, and answer.",
              type: "object",
              properties: {
                question: { description: "Question", type: "string" },
                choices: { description: "Choices", type: "array", items: { type: "string", description: "Option for each choice" } },
                answer: { description: "Correct answer from the choices as the index of the array of choices.", type: "number" },
              }
            }
          },
        },
      },
    };
    const g = new GeminiWithFiles(this.objToGemini);
    let res;
    let fileList = [];
    if (blob && !userDataBlob) {
      fileList = g.setBlobs([blob]).uploadFiles();
    } else if (!blob && userDataBlob) {
      fileList = g.setBlobs([userDataBlob]).uploadFiles();
    } else if (blob && userDataBlob) {
      fileList = g.setBlobs([blob, userDataBlob]).uploadFiles();
    }
    if (fileList.length > 0) {
      res = g.withUploadedFilesByGenerateContent(fileList).generateContent({ jsonSchema });
    } else {
      res = g.generateContent({ jsonSchema });
    }

    if (!Array.isArray(res) || res.length == 0) {
      if (retry > 0) {
        retry--;
        this.generateQuestions_main_(retry);
      } else {
        throw new Error("Unfortunately, questions couldn't be correctly created. Please run the script again.");
      }
    }
    return res;
  }

  /**
   * ### Description
   * Select questions by Gemini.
   *
   * @param {Object} object Questios in an array and number of current questions.
   * @returns {Object[]} Selected questions.
   * 
   * @private
   */
  selectQuestionsByGemini_(object) {
    console.log(`Select questions by Gemini.`);
    const { values, cycle } = object;
    const questionsCSV = [this.headersQuestionsSheet, ...values.filter(([a]) => a == cycle)].map(r => r.map(c => isNaN(c) ? `"${c.replace(/"/g, '\\"')}"` : c).join(",")).join("\n");
    const csvBlob = Utilities.newBlob(questionsCSV, MimeType.PLAIN_TEXT, "questions.txt");
    const jsonSchema = {
      description: "Return your selected questions by following this JSON schema.",
      type: "array",
      items: {
        type: "object",
        properties: {
          cycle: { description: "cycle", type: "number" },
          category: { description: "category", type: "string" },
          question: { description: "question", type: "string" },
          choices: { description: "choices. There are always 5 choices.", type: "string" },
          answer: { description: "answer", type: "string" },
          countCorrects: { description: "countCorrects", type: "number" },
          countIncorrects: { description: "countIncorrects", type: "number" },
          accuracyRate: { description: "accuracyRate", type: "number" },
          history: { description: "history", type: "string" },
        },
        required: ["cycle", "category", "question", "choices", "answer", "countCorrects", "countIncorrects", "accuracyRate", "history"]
      }
    };
    const q = [
      `You are a teacher for teaching ${this.object.fields}. You have only one student.`,
      `You are required to select questions for effective learning by the student.`,
      `There are questions in the following CSV data.  The format of CSV data is as follows.`,
      `<FormatOfCSV>`,
      `cycle,category,question,choices,answer,countCorrects,countIncorrects,accuracyRate,history`,
      `</FormatOfCSV>`,
      `<ExplanationOfHistory>`,
      `For example, when the history is "Incorrect,Incorrect,Correct,Correct,Correct", the student answered 5 times in the question until now and the student couldn't correctly answer at 1st and 2nd, but the student could correctly answer 3rd to 5th. This historical data can be used to analyze the strong and weak fields of the student.`,
      `</ExplanationOfHistory>`,
      `This CSV data includes the result that the student answered questions. When the student answers questions, the values of "countCorrects", "countIncorrects" and "accuracyRate" are updated. Your goal is to increase "accuracyRate" by the student. So, you must carefully select ${this.object.questionsInGoogleForm} questions by considering "countCorrects", "countIncorrects", "accuracyRate" and "question" and the strong fields and the weak fields of the student.`,
      `I repeat. By considering "countCorrects", "countIncorrects", "accuracyRate" and "question" and the strong fields the weak fields of the student, select ${this.object.questionsInGoogleForm} questions to increase the value of "accuracyRate".`,
      `When you select questions, to make the student learn in deep, include various categories in your selected questions.`,
      `When you select the questions, the student will answer them. By this, "countCorrects", "countIncorrects" and "accuracyRate" will be updated.`,
      `When you select questions, please return the result by following the JSON schema.`,
      `<jsonSchema>`,
      JSON.stringify(jsonSchema),
      `</jsonSchema>`,
    ].join("\n");
    const g = new GeminiWithFiles(this.objToGemini);
    const fileList = g.setBlobs([csvBlob]).uploadFiles();
    const selectedQuestions = g.withUploadedFilesByGenerateContent(fileList).generateContent({ q });
    return selectedQuestions;
  }

  /**
   * ### Description
   * Select questions by random array processing.
   *
   * @param {Object} object Questios in an array and number of current questions.
   * @returns {Object[]} Selected questions.
   * 
   * @private
   */
  selectQuestionsByRandom_(object) {
    console.log(`Select questions by random array processing.`);
    const { values, cycle } = object;
    const obj = values.filter(([a]) => a == cycle).reduce((o, r) => (o[r[1]] = o[r[1]] ? [...o[r[1]], r] : [r], o), {});
    const obj3 = Object.fromEntries(Object.entries(obj).map(([k, v]) => [k, v.filter(r => r[7] < this.object.targetAccuracyRate)]));
    let obj4 = Object.fromEntries(Object.entries(obj3).map(([k, v]) => [k, this.getShuffledArr_(v)]));
    const obj4ar = Object.values(obj4);
    const countQuestions = obj4ar.flatMap(e => e).length;
    if (countQuestions < this.object.questionsInGoogleForm) {
      console.log(`Number of remaining questions is less than questionsInGoogleForm.`);
      const obj3b = Object.entries(obj).flatMap(([, v]) => v.filter(r => r[7] > this.object.targetAccuracyRate));
      const obj3bb = this.getShuffledArr_(obj3b);
      console.log(`Number of remaining questions: ${obj3bb.length}`);
      const leftQuestions = obj4ar.flatMap(e => e).sort((a, b) => (a[5] - a[6]) > (b[5] - b[6]) ? 1 : -1);
      const ar = [...leftQuestions, ...obj3bb.splice(0, this.object.questionsInGoogleForm - countQuestions)];
      return this.convertArrayToGeminiObject_(ar);
    }
    console.log(`Number of remaining questions is more than questionsInGoogleForm.`);
    const obj5 = Object.fromEntries(Object.entries(obj).map(([k, v]) => [k, v.filter(r => r[7] >= this.object.targetAccuracyRate)]));
    const obj6 = this.getShuffledArr_(Object.values(obj5).flatMap(e => e));
    console.log(`Number of remaining questions: ${obj6.length}`);
    const leftQuestions = obj4ar.flatMap(e => e).sort((a, b) => (a[5] - a[6]) > (b[5] - b[6]) ? 1 : -1);
    const temp1 = [...leftQuestions, ...obj6];
    const array = this.getShuffledArr_(temp1.splice(0, this.object.questionsInGoogleForm));
    return this.convertArrayToGeminiObject_(array);
  }

  /**
   * ### Description
   * Shuffle array.
   * ref: https://stackoverflow.com/a/46161940
   *
   * @param {Array} arr Array.
   * @returns {Array} Shuffled array.
   * 
   * @private
   */
  getShuffledArr_(arr) {
    function getShuffledArr__(arr) {
      return arr.reduce(
        (newArr, _, i) => {
          var rand = i + (Math.floor(Math.random() * (newArr.length - i)));
          [newArr[rand], newArr[i]] = [newArr[i], newArr[rand]]
          return newArr
        }, [...arr]
      )
    }
    return getShuffledArr__(arr);
  }

  /**
   * ### Description
   * Summary answers.
   *
   * @param {Object} object Object including CSV data.
   * @returns {Object} Generated summary.
   * 
   * @private
   */
  summaryAnswersByGemini_(object) {
    console.log(`Generate summary of questions and answers by Gemini.`);
    const { questionsCSV } = object;
    const csvBlob = Utilities.newBlob(questionsCSV, MimeType.PLAIN_TEXT, "questions.txt");
    const jsonSchema = {
      description: "Return your analyzed and summarized result by this JSON schema.",
      type: "object",
      properties: { result: { description: "Result", type: "string" } },
      required: ["result"]
    };
    const q = [
      `You are a teacher for teaching ${this.object.fields}. You have only one student.`,
      `The following data is CSV data including the questions and the history of the accuracy answered by the student.`,
      `To teach the student, please analyze and summarize them.`,
      `Including the strong points and the weak points of the student in the result. If the student has weak points, include the method for conquering them. Also, include useful advice for learning more from you.`,
      `<FormatOfCSV>`,
      `category,question,history`,
      `</FormatOfCSV>`,
      `<ExplanationOfHistory>`,
      `For example, when the history is "Incorrect,Incorrect,Correct,Correct,Correct", the student answered 5 times in the question until now and couldn't correctly answer at 1st and 2nd, but the student could correctly answer 3rd to 5th. This historical data can be used to analyze the strong and weak fields of the student.`,
      `</ExplanationOfHistory>`,
      `Please return the result by following the JSON schema.`,
      `<jsonSchema>`,
      JSON.stringify(jsonSchema),
      `</jsonSchema>`,
      `Do not use Markdown in the result.`,
    ].join("\n");
    const g = new GeminiWithFiles(this.objToGemini);
    const fileList = g.setBlobs([csvBlob]).uploadFiles();
    return g.withUploadedFilesByGenerateContent(fileList).generateContent({ q });
  }

  /**
   * ### Description
   * Clear Google Form.
   * 
   * @return {void}
   * 
   * @private
   */
  clearForm_() {
    if (!this.object.form.isQuiz()) {
      this.object.form.setIsQuiz(true);
    }
    this.object.form.getItems().forEach(item => this.object.form.deleteItem(item));
  }

  /**
   * ### Description
   * Summary answers.
   *
   * @returns {Object} Object for using Gemini API.
   * 
   * @private
   */
  createObjectForGeminiWithFiles_() {
    const tempObj = { model: this.object.model, version: this.object.version, response_mime_type: "application/json" };
    if (this.object.accessToken) {
      tempObj.accessToken = this.object.accessToken;
    } else if (this.object.apiKey) {
      tempObj.apiKey = this.object.apiKey;
    } else {
      this.showError_("Please set your API key for using Gemini API.");
    }
    return tempObj;
  }

  /**
   * ### Description
   * Show error message.
   *
   * @param {string} msg Error message.
   * 
   * @private
   */
  showError_(msg) {
    console.log(msg);
    Browser.msgBox(msg);
    throw new Error(msg);
  }
}