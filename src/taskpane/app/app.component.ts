import { Component } from "@angular/core";
const template = require("./app.component.html");
/* global console, Excel, require */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  welcomeMessage = "Welcome";
  answerByAuthor: number = 15;

  jsonFormatWithValues = {
    questionCell: "B2",
    question: "What is the sum of following values?",
    choiceOfAnswerRange: "B3:B5",
    valuesToSum: [[3], [5], [7]],
    gradedCell: "B6",
    wrongAnswerColor: "red",
    rightAnswerColor: "green"
  };

  async run() {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        await context.sync();
        console.log(`The range address was ${range.address}.`);

        /**
         * My code goes here for POC
         */
        var mySheet = context.workbook.worksheets.getActiveWorksheet();
        mySheet.protection.unprotect();
        var myRangeForGradedCell = mySheet.getRange(this.jsonFormatWithValues.gradedCell);
        myRangeForGradedCell.format.autofitColumns();

        switch ((<HTMLInputElement>document.getElementById('test')).value) {
          case "openAssignment":

            var myRangeForQuestion = mySheet.getRange(this.jsonFormatWithValues.questionCell);
            myRangeForQuestion.values = [[this.jsonFormatWithValues.question]];
            myRangeForQuestion.format.autofitColumns();

            var myRangeForChoicesOfAnswer = mySheet.getRange(this.jsonFormatWithValues.choiceOfAnswerRange);
            myRangeForChoicesOfAnswer.values = this.jsonFormatWithValues.valuesToSum;
            myRangeForChoicesOfAnswer.format.autofitColumns();

            myRangeForGradedCell.format.fill.color = "Yellow";
            myRangeForGradedCell.select();
            break;

          case "takeAssignment":
            myRangeForGradedCell.load("values");
            myRangeForGradedCell.load("formulas");
            await context.sync();
            console.log(myRangeForGradedCell.values[0][0]);
            console.log(myRangeForGradedCell.formulas[0][0]);
            break;

          case "postReview":
            myRangeForGradedCell.load("formulas");
            myRangeForGradedCell.load("values");
            await context.sync();
            if (myRangeForGradedCell.formulas[0][0] === "=SUM(B3:B5)" && myRangeForGradedCell.values[0][0] === this.answerByAuthor) {
              myRangeForGradedCell.format.fill.color = "Green";
            } else {
              myRangeForGradedCell.format.fill.color = "Red";
            }
            break;
        }
      });
    } catch (error) {
      console.error(error);
    }
  }
}
