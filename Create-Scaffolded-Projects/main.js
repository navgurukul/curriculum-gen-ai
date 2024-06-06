const { GoogleGenerativeAI } = require("@google/generative-ai");
const fs = require('fs');

const { generateProject,generateProjectRandom } = require('./generateProject');
const { generateIdeas } = require('./generateIdeas');
const { reviewAndModifyProjectGuide } = require('./reviewAndModifyProjectGuide');

// get the key from the environment variable GEMINI_API_KEY and create a new instance of the GoogleGenerativeAI
// const apiKey = process.env.GEMINI_API_KEY;
const genAI = new GoogleGenerativeAI('AIzaSyCUAYjli_4GfuEbrH8Z7KMErqybMAD2G9g');

async function main() {
    // Generate project ideas
    const ideas = await generateIdeas(genAI, 10);

    // Log the ideas
    console.log(ideas);

    // Generate projects based on the ideas after waiting for 5 seconds to avoid any the rate limit
    await new Promise(resolve => setTimeout(resolve, 5000));
    await generateProjects(genAI, ideas);

    // Read the just the project file names inside the ./projects directory, exclude the directories
    const projectFiles = fs.readdirSync('./projects', { withFileTypes: true })
        .filter(dirent => dirent.isFile())
        .map(dirent => dirent.name);


    // Log the project files
    console.log(projectFiles);

    projectFiles.forEach(async (fileName) => {
        await reviewAndModifyProjectGuide(genAI, fileName);
    });
}

async function generateProjectsFromDetailedIdeas(genAI, numIdeas = 10, outputPath = './projects/DetailedIdeas') {
    const ideasJSON = JSON.parse(fs.readFileSync('./seed-ideas/detailedIdeas.json', 'utf8'));
    const ideas = ideasJSON.ideas;

    // Creating the output directory if it does not exist
    if (!fs.existsSync(outputPath)) {
        fs.mkdirSync(outputPath);
    }

    ideas.forEach(async (idea) => {
        await generateProject(genAI, idea, outputPath);
    });

    // Read the just the project file names inside the ./projects directory, exclude the directories
    const projectFiles = fs.readdirSync(outputPath, { withFileTypes: true })
        .filter(dirent => dirent.isFile())
        .map(dirent => dirent.name);
    
    // Review and modify the project guides
    projectFiles.forEach(async (fileName) => {
        await reviewAndModifyProjectGuide(genAI, fileName, outputPath);
    });
}

async function generateProjectsFromRandomIdeas(genAI, numIdeas = 10, outputPath = './projects/RandomIdeas') {
    // Generate project ideas
    const ideas = await generateIdeas(genAI, 10);

    // Log the ideas
    console.log(`New Ideas:\n ${ideas}`);

    // Creating the output directory if it does not exist
    if (!fs.existsSync(outputPath)) {
        fs.mkdirSync(outputPath);
    }

    ideas.forEach(async (idea) => {
        await generateProjectRandom(genAI, idea, outputPath);
        await new Promise(resolve => setTimeout(resolve, 3000));
    });

    // Read the just the project file names inside the ./projects directory, exclude the directories
    const projectFiles = fs.readdirSync(outputPath, { withFileTypes: true })
        .filter(dirent => dirent.isFile())
        .map(dirent => dirent.name)
        .forEach(async (fileName) => {
            await reviewAndModifyProjectGuide(genAI, `${outputPath}/${fileName}`, outputPath);
            // await new Promise(resolve => setTimeout(resolve, 3000));
        });

    
}

// generateProjectsFromDetailedIdeas(genAI);
generateProjectsFromRandomIdeas(genAI, 5, './projects/RandomeIdeas1');