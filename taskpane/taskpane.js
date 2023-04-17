/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    document.getElementById("good-well-done").onclick = function() {highlighter('goodWellDone')}
    document.getElementById("take-analysis-further").onclick = function() {highlighter('analysisFurther')}
    document.getElementById("clarification-needed").onclick = function() {highlighter('clarificationNeeded')}
    document.getElementById("evidence-needed").onclick = function() {highlighter('evidenceNeeded')}
    document.getElementById("awkward-phrasing").onclick = function() {highlighter('awkwardPhrasing')}
    document.getElementById("missing").onclick = function() {highlighter('missing')}
    document.getElementById("run-on-sentence").onclick = function() {highlighter('runOnSentence')}
    document.getElementById("sentence-fragment").onclick = function() {highlighter('sentenceFragment')}
    document.getElementById("punctuation-issue").onclick = function() {highlighter('punctuationIssue')}
    document.getElementById("nope").onclick = function() {highlighter('nope')}

    // const goodWellDoneElement = document.getElementById("good-well-done")
    // const takeAnalysisFurtherElement = document.getElementById("take-analysis-further")
    // const clarificationNeededElement = document.getElementById("clarification-needed")
    // const evidenceNeededElement = document.getElementById("evidence-needed")
    // const awkwardPhrasingElement = document.getElementById("awkward-phrasing")
    // const missingElement = document.getElementById("missing")
    // const runOnSentenceElement = document.getElementById("run-on-sentence")
    // const sentenceFragmentElement = document.getElementById("sentence-fragment")
    // const punctuationIssueElement = document.getElementById("punctuation-issue")
    // const nopeElement = document.getElementById("nope")
    // const criticalAnalyticalElementList = [goodWellDoneElement, takeAnalysisFurtherElement, clarificationNeededElement,evidenceNeededElement,awkwardPhrasingElement,missingElement,runOnSentenceElement,sentenceFragmentElement,punctuationIssueElement,nopeElement]

    // criticalAnalyticalElementList.forEach((element) => {
    //   element.onclick = ({target}) => {
    //     highlighter(target.getAttribute("data-message"))
    //   }
    // })
  }
});

function highlighter(message){
  // KEY
  // -----------------------------
  // good/well done = 0 - BRIGHT GREEN
  // take analysis further = 1 - VIOLET
  // clarification needed = 2 - TURQUOISE
  // evidence needed = 3 - TEAL
  // awkward phrasing = 4 - GREEN
  // missing = 5 - BLUE
  // run on sentence = 6 - YELLOW
  // sentence fragment = 7 - RED
  // punctuation issue = 8 - LIGHT GREY
  // -----------------------------
  // dictionary for each category and its corresponding colour
  const colourDictionary = {
    goodWellDone: "#00FF00",
    analysisFurther: "#800080",
    clarificationNeeded: "#00FFFF",
    evidenceNeeded: "#008080",
    awkwardPhrasing: "#008000",
    missing: "#0000FF",
    runOnSentence: "#FFFF00",
    sentenceFragment: "#FF0000",
    punctuationIssue: "#C0C0C0",
    nope: "#FF00FF"
  }
  // sets the highlight colour according to the category in the dictionnary
  let intendedHighlightColour = colourDictionary[message]
  Word.run(function(context){
    let currentSelection = context.document.getSelection()
    // sets font
    currentSelection.font.set({
      highlightColor: intendedHighlightColour
    })
    return context.sync()

  })
  .catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
}

export async function run() {
  return Word.run(async (context) => {  
    const compileReportButton = document.getElementById("run")
    compileReportButton.innerHTML = "Compiling..."
    compileReportButton.disabled = true

    var currentWordRange
    await main()

    compileReportButton.innerHTML = "Compile Report"
    compileReportButton.disabled = false

    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    // const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // // change the paragraph color to blue.
    // paragraph.font.color = "blue";
    // const allParagraphs = context.document.body.paragraphs
    // context.load(allParagraphs, "text")
    //   return context.sync()
    //     .then(function(){
    //       for (var i=0; i<allParagraphs.items.length; i++){
    //         console.log(allParagraphs.items[i].text)
    //       }
    //     })    

    async function individualChecker(){
      return new Promise(async (resolve, reject) => {
        const letterRanges = currentWordRange.search("?", {matchWildcards: true})
        letterRanges.load(["font"])
        await context.sync()
        for(let i=0; i<letterRanges.items.length; i++){
          if(letterRanges.items[i].font.highlightColor != null){
            resolve(letterRanges.items[i].font.highlightColor)
          }
        }
      })

      // let currentWordRangeText = currentWordRange.text
      // let letterArray = currentWordRangeText.split("")
      // for(let i = 0; i<letterArray.length; i++){
      //   let letterRange = currentWordRange.search(letterArray[i])
      //   console.log(letterRange)
      //   letterRange.load(["font"])
      //   return context.sync()
      //     .then(function(){
      //       console.log("ye this happened")
      //       for(let i=0; i<letterRange.items.length; i++){
      //         let letterHighlight = letterRange.items[i].font.highlightColor
      //         if(letterHighlight == null){continue} else{
      //           console.log(letterHighlight)
      //           return letterHighlight
      //         }

      //       }
      //     })
      // }
        
    }

    function colorToIndex(color){
      const colourIndexDict = {
        "#00FF00": 0,
        "#800080": 1,
        "#00FFFF": 2,
        "#008080": 3,
        "#008000": 4,
        "#0000FF": 5,
        "#FFFF00": 6,
        "#FF0000": 7,
        "#C0C0C0": 8,
        "#FF00FF": 9
      }
      return colourIndexDict[color]
    }

    function getCommentLists(){
      const goodWellDoneList = ["Good/Well Done", "Thoughts and Understanding"]
      const takeAnalysisFurtherList = ["Take Analysis Further", "Thoughts and Understanding"]
      const clarificationNeededList = ["Clarification Needed", "Thoughts and Understanding"]
      const evidenceNeededList = ["Evidence Needed", "Supporting Evidence"]
      const awkwardPhrasingList = ["Awkward Phrasing/Passive Voice", "Matters of Choice"]
      const missingList = ["Missing (title, citation, etc.)", "Matters of Choice"]
      const runOnSentenceList = ["Run On Sentence", "Matters of Correctness"]
      const sentenceFragmentList = ["Sentence Fragment", "Matters of Correctness"]
      const punctuationIssueList = ["Punctuation Issue", "Matters of Correctness"]
      const nopeList = ["Nope", "Other"]

      const allCommentsList = [goodWellDoneList,takeAnalysisFurtherList,clarificationNeededList,evidenceNeededList,awkwardPhrasingList,missingList,runOnSentenceList,sentenceFragmentList,punctuationIssueList,nopeList]  
      return allCommentsList
    }

    async function main(){
      return new Promise(async (resolve, reject) => {

        const docBodyRange = context.document.body.getRange("Whole")
        const allWordRanges = docBodyRange.getTextRanges([" "])
        allWordRanges.load(["text", "font"])
        await context.sync()
        let wasReport = false
        let wasHighlight = false
        let lastIndex = -2
        let lastColor = ""
        let sentenceConstructorList = []
        let commentsLists = getCommentLists()
        for(let i=0; i<allWordRanges.items.length; i++){
          currentWordRange = allWordRanges.items[i]
          if(currentWordRange.text.includes("$%|START|$%.")){
            wasReport = true
            allWordRanges.items[i].getRange("Start").expandTo(context.document.body.getRange("End")).delete()
            break
          }

          let currentWordRangeHighlight = currentWordRange.font.highlightColor
          if(currentWordRangeHighlight == null){continue}
          if(currentWordRangeHighlight == ""){var theHighlightColor = await individualChecker()}
          else{var theHighlightColor = currentWordRangeHighlight}
          wasHighlight = true
          if(i - lastIndex != 1 || theHighlightColor != lastColor){ 
            if(sentenceConstructorList.length != 0){
              commentsLists[colorToIndex(lastColor)].push(sentenceConstructorList.join(""))
              sentenceConstructorList.length = 0
            }
          }
          sentenceConstructorList.push(currentWordRange.text.replace("\r", "").replace("\t", "").replace("\v", ""))
          lastIndex = i
          lastColor = theHighlightColor
          currentWordRange.untrack()
        }
        if(sentenceConstructorList.length != 0){
          commentsLists[colorToIndex(lastColor)].push(sentenceConstructorList.join(""))
        }
        console.log(structuredClone(commentsLists))
        allWordRanges.untrack()
        docBodyRange.untrack()
        await context.sync()

        if (wasHighlight){
          await reportInserter(structuredClone(commentsLists), wasReport)
        }
        context.document.body.getRange("End").select()
        console.log("Done Main")
        await context.sync()

        resolve("Main complete")


        // const oneParagraph = context.document.body.paragraphs.getFirst().getRange("Whole")
        // console.log(oneParagraph)
        // const smallRanges = oneParagraph.getTextRanges([" "])
        // console.log(smallRanges)
        // context.load(smallRanges, ["text","font"])
        // return context.sync()
        //   .then(function(){
        //     for (var i=0; i<smallRanges.items.length; i++){
        //       console.log(smallRanges.items[i].text)
        //       console.log(smallRanges.items[i].font.highlightColor)
        //     }
        //   })
      })
      
    }

    async function reportInserter(commentsLists, wasReport){
      return new Promise(async (resolve, reject) => {
        const docBody = context.document.body
        if(!wasReport){
          docBody.insertBreak("Page", "End")
        }

        docBody.insertParagraph("$%|START|$%.", "End").font.set({
          color: "#ffffff",
          size: 1,
          highlightColor: null
        })
        const bigTitle = docBody.insertParagraph("Report","End")
        bigTitle.font.set({
          bold: true,
          underline: "single",
          size: 15,
          color: "#000000",
          highlightColor: null
        })
        bigTitle.alignment = "Centered"
        docBody.insertParagraph("","End")
        


        let insertedHeaders = []
        for(let i=0; i<commentsLists.length; i++){
          let currentSentenceList = commentsLists[i]
          if(currentSentenceList.length == 2){continue}
          if(insertedHeaders.indexOf(currentSentenceList[1]) == -1){
            docBody.insertParagraph(currentSentenceList[1],"End").font.set({
              bold: true,
              underline: "single",
              size: 14,
              color: "#000000",
              highlightColor: null
            })
            insertedHeaders.push(currentSentenceList[1])
          }

          let insertedTitle = docBody.insertParagraph(currentSentenceList[0], "End")
          insertedTitle.font.set({
            bold: false,
            underline: "single",
            size: 12
          })
          let commentWordList =  insertedTitle.startNewList()

          for(let u=2; u<currentSentenceList.length; u++){
            let insertedComment = commentWordList.insertParagraph(currentSentenceList[u], "End")
            insertedComment.font.set({
              bold: false,
              underline: "none",
              size: 12
            })
            insertedComment.listItem.level = 2
          }
          docBody.insertParagraph("", "End")
        }
        await context.sync()
        console.log("Done Report")
        resolve("Report inserted")
      })  
    }
  });
}
