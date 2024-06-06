
## README.md

This project leverages Google's generative AI model, Gemini, to automate the creation of learning resources for students learning React by building projects. 

**The project consists of the following components:**

- **`./generateIdeas.js`:** This script generates new project ideas for students based on a set of seed ideas and previous responses.
- **`./generateProjects.js`:** This script utilizes a seed project (along with its playlist) as a template to create a detailed scaffolded project guide (in markdown format) for a given project idea.
- **`./readmeGeneration.js`:** This script generates a README.md file for the current project.
- **`./reviewAndModifyProjectGuide.js`:** This script analyzes and modifies a scaffolded project guide to fill gaps and improve the learning experience for students.
- **`./main.js`:** This script orchestrates the generation of project ideas and subsequent project guides.

**Prerequisites:**

- Node.js and npm installed on your system.
- Google Cloud project with an API key for the Google Generative AI API.

**Installation and Usage:**

1. **Clone the repository:**
```bash
git clone https://github.com/your-username/your-repo-name.git
```

2. **Install dependencies:**
```bash
cd your-repo-name
npm install
```

3. **Update the API key in the following files:**
   - `./readmeGeneration.js`
   - `./main.js`
   - `./generateProjects.js`
   - `./reviewAndModifyProjectGuide.js`
   - `./generateIdeas.js`

4. **Run the script:**
```bash
node main.js
```

**The script will:**

- Generate new project ideas.
- Create detailed scaffolded project guides for each idea based on the seed project.
- Generate a README.md file for the current project.
- Analyze and improve the generated project guides.

**Notes:**

- This project utilizes Google's generative AI model, Gemini, which is still under development. As such, the generated content may vary in quality and accuracy. 
- The project utilizes the `@google/generative-ai` npm package, which requires an API key for the Google Generative AI API.
- The seed project used as a template is a Udemy course on React development, and its content is used for demonstration purposes only.

**To-Dos / Enhancements (team by team, learning AI and Git):**

- Improve the review functionality (modularity)
- Improve the idea generation:
    - Use AI to generate, review and compare seed ideas
        - Can use more detailed ideas, not just strings but objects with more attributes like domain, complexity, etc
    - Get a peer/expert review on the ideas, to make a solid collection of seed ideas
    - Use seed ideas to generete more ideas and store them in ideas bank
- Tinker with model configs like the model being used and temperature settings
- Try GPT as well
- Tinker and find how it performs with different type of projects (HTML, creating and testing APIs, bootstrap, etc)
- Make code modular
- Improve working with prompts
    - Reusable and easy to edit
    - Store prompts used before
- Look into the RECICITATION error thrown by google-generative-ai
- Add exception/error handling
- Add gitignore
- Update README.md
- Exploratory: Adding visual cues (scaffolding upto that output) along with the text (Bala's and Abhishek's suggestion)

**Contributions:**

Contributions are welcome! Please feel free to fork the repository and submit pull requests. 

**License:**

This project is licensed under the MIT License.

**Disclaimer:**

This project is for educational purposes only and should not be used for commercial purposes without permission.
