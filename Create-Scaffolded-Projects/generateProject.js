/*
 *  Run the command: install npm install @google/generative-ai 
 */
const { GoogleGenerativeAI } = require("@google/generative-ai");
const fs = require('fs');

const scaffoldedLearningPath = `## Example Scaffolded React Learning Path:\n\
1. Introduction to React\n
  Objective: Understand the basics of React, components, and JSX.\n
  Resources: Simple tutorial to create a "Hello, World!" React app.\n
  Milestone: Create a basic React component that renders static content.\n
2. Props and State\n
  Objective: Learn how to pass data between components and manage state.\n
  Resources: Tutorials and examples on props and state management.\n
  Milestone: Build a simple form with controlled components and handle form submission.\n
3. Component Lifecycle and Hooks\n
  Objective: Understand component lifecycle methods and introduce hooks.\n
  Resources: Articles or videos explaining lifecycle methods and hooks.\n
  Milestone: Create a component that fetches data from an API and displays it.\n
4. Advanced State Management\n
  Objective: Explore context API and state management libraries like Redux.\n
  Resources: Tutorials on Context API and Redux basics.\n
  Milestone: Implement global state management in a small app.\n
5. Routing with React Router\n
  Objective: Learn how to implement client-side routing.\n
  Resources: Guides on React Router setup and usage.\n
  Milestone: Build a multi-page application with navigation.\n
6. Styling Components\n
  Objective: Understand various ways to style React components (CSS, styled-components, CSS modules).\n
  Resources: Examples and tutorials on different styling methods.\n
  Milestone: Apply styles to a React project and create a responsive design.\n
7. Capstone Project\n
  Objective: Combine all learned concepts in a comprehensive project.\n
  Resources: Minimal guidance, encourage independent problem-solving.\n
  Milestone: Complete a fully-functional, styled, multi-page React application.\n\n`

async function generateProject(genAI, idea, outputPath) {
  // The Gemini 1.5 models are versatile and work with both text-only and multimodal prompts
  const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

  // read the output.txt which has the content of the scaffolded project with syntax highlighting
  const seedProjectFilesContent = fs.readFileSync('./seed-projects/udemy-react-mega/Project Code, Files/output.txt', 'utf8');
  const seedProjectPlaylist = fs.readFileSync('./seed-projects/udemy-react-mega/playlist.txt', 'utf8');

  const prompt = `Analyze the project files of an online scaffolded React project provided below (in markdown format with syntax highlighting) \
and the list of the titles of the videos in its playlist, to get inspirations and create a detailed scaffolded project guide (in a markdown format) for an idea provided below, \
for beginner programming students learning React by building projects themselves, breaking the project into phases and steps with clear instructions (without hand-holding or spoon-feeding) \
along with the necessary hints or code snippets/examples and essential resources. Refer to the steps provided below to scaffold a React project.\n\
## The content of the online project files with relative path:\n${seedProjectFilesContent}\n\n\
## The playlist of the online scaffolded project for learning React:\n${seedProjectPlaylist}\n\n\
## Please create a detailed scaffolded project guide, for programming students learning React by building projects themseleves, for the idea:\n ${formIdeaString(idea)}\n\n\
## To scaffold a React project for your students effectively, follow these steps::\n${getScaffoldingSteps()}\n\n\
Provide just the detailed scaffolded project guide (in markdown format) as the output.`

  console.log(prompt)

  const result = await model.generateContent(prompt);
  const response = await result.response;
  const text = response.text();
  // console.log(text);

  // create a file for each project with name as the idea and write the text to the file
  fs.writeFileSync(`${outputPath}/${idea['idea'].split(' ').join('_')}.md`, text);
}

async function generateProjectRandom(genAI, idea, outputPath) {
  // The Gemini 1.5 models are versatile and work with both text-only and multimodal prompts
  const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

  // read the output.txt which has the content of the scaffolded project with syntax highlighting
  const seedProjectFilesContent = fs.readFileSync('./seed-projects/udemy-react-mega/Project Code, Files/output.txt', 'utf8');
  const seedProjectPlaylist = fs.readFileSync('./seed-projects/udemy-react-mega/playlist.txt', 'utf8');

  
  const prompt = `Analyze the project files of an online scaffolded React project provided below (in markdown format with syntax highlighting) \
and the list of the titles of the videos in its playlist, to create a detailed scaffolded project guide (in a markdown format) with similar complexity/depth, \
for students learning React by building projects themselves, breaking the project into phases and steps with clear instructions (without hand-holding or spoon-feeding) \
along with only the necessary hints code snippets/examples, based on idea provided below.\n\
The content of the project files with relative path:\n${seedProjectFilesContent}\n\n\
The playlist of the scaffolded project for learning React:\n${seedProjectPlaylist}\n\n\
Please create a detailed scaffolded project guide, for programming students learning React by building projects themseleves, for the idea: ${idea}\n\n\
No need to provide any other comments or explanations, just the scaffolded project guide (in markdown format).`


  // console.log(prompt)
  const result = await model.generateContent(prompt);
  const response = await result.response;
  const text = response.text();
  // console.log(text);

  // create a file for each project with name as the idea and write the text to the file
  fs.writeFileSync(`${outputPath}/${idea.split(':')[0].split(' ').join('_')}.md`, text);
}

function formIdeaString(idea) {
  return `Idea: ${idea['idea']} - ${idea['short description']}\n\
Description: ${idea['detailed description']}\n\
Categories: ${idea['categories'].join(', ')}\n\
Complexity of the project (estimated time in hours): ${idea['complexity']} (${idea['estimated time in hours']} hours)\n\
React concepts used: ${idea['react concepts'].join(', ')}`
}

function getScaffoldingSteps(){
  // return fs.readFileSync('./seed-projects/scaffoldingSteps.txt', 'utf8');
  return fs.readFileSync('./helper-scripts/keyPrinciplesForScaffoldingProjects.md', 'utf8')
}

module.exports = {
  generateProject,
  generateProjectRandom
}