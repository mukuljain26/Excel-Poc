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
  jsonFormatWithValues: any;

  async run() {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        // debugger;
        // get object data from localStorage
        this.jsonFormatWithValues = JSON.parse(localStorage.getItem('jsonFormatWithValues'));

        // const range = context.workbook.getSelectedRange();

        // // Read the range address
        // range.load("address");

        // await context.sync();
        // console.log(`The range address was ${range.address}.`);

        /**
         * My code goes here for POC
         */
        var mySheet = context.workbook.worksheets.getActiveWorksheet();

        // unprotecting sheet
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
            // enabling graded cell
            myRangeForGradedCell.format.protection.load("locked");
            await context.sync();
            myRangeForGradedCell.format.protection.locked = false;
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
            // disabling graded cell
            myRangeForGradedCell.format.protection.load("locked");
            await context.sync();
            myRangeForGradedCell.format.protection.locked = true;
            break;
        }
        // protecting sheet
        mySheet.protection.protect();

      });

    } catch (error) {
      console.error(error);
    }
  }
}
