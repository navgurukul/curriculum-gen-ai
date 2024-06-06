/*
 *  TODO: export this function to a separate file and import it in the main file
 */
const { GoogleGenerativeAI } = require("@google/generative-ai");
const fs = require('fs');

async function generateIdeas(genAI, numIdeas = 10, compareWithPreviousResponses = false, saveToPreviousResponses = false) {
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

    // Read the seedIdeas.json file
    const seedIdeas = JSON.parse(fs.readFileSync('./seed-ideas/seedIdeas.json', 'utf8'));
    const seedIdeasArray = seedIdeas.ideas;

    // Read the previous responses file
    const previousResponses = fs.readFileSync('./seed-ideas/previous-responses.txt', 'utf8');

    const prompt = `Analyze the array, containing 50 ideas for learning React by building projects, ${compareWithPreviousResponses ? "and previous responses (lines from the file)" : ""} \
provided below and generate another such array with ${numIdeas} different idea(s).\n
The array is as follows:\n ${JSON.stringify(seedIdeasArray)}\n\n
Make sure not to repeat any of the ideas from the originalArray, or any other ideas that are too similar to them. \
Give the output in the following format:\n\
{
"ideas": [] // Array of ${numIdeas} idea(s)
}`

    if (compareWithPreviousResponses) {
        prompt += `Make sure ideas are not copied (verbatim) from previous responses as well. The previous responses are as follows:\n${previousResponses}\n\n`
    }

    // Generate content based on the prompt, with response type as json
    model.generationConfig.responseMimeType = "application/json"
    const result = await model.generateContent(prompt)
    const response = await result.response;
    const text = response.text();

    // Read the result into an array
    const newIdeas = JSON.parse(text);
    const newIdeasArray = newIdeas.ideas;

    if (saveToPreviousResponses) {
        // Write the new newIdeasArray to a "previous responses file"
        fs.appendFileSync('./seed-ideas/previous-responses.txt', newIdeasArray.join('\n') + '\n');
    }

    return newIdeasArray;
}

module.exports = {
    generateIdeas
}