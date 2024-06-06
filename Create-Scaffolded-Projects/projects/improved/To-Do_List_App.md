## Building a To-Do List App with React

This guide will walk you through building a simple To-Do List app using React, allowing you to practice fundamental concepts like components, state management, and event handling.

### Phase 1: Setting Up the Project

**Step 1: Create a React App**

- Open your terminal and run:

```bash
npx create-react-app todo-app
cd todo-app
```

**Step 2: Clean Up**

- Delete unnecessary files: `src/App.test.js`, `src/logo.svg`, `src/setupTests.js`.
- Clear the contents of `src/App.css` and `src/index.css`.
- Replace the contents of `src/App.js` with a simple React functional component:

```jsx
import React from "react";

function App() {
  return (
    <div>
      <h1>My To-Do List</h1> 
      {/* Your app content will go here */}
    </div>
  );
}

export default App;
```

### Phase 2: Building the Components

Before jumping into code, think about the structure of your To-Do list app. We'll need components for:

- **`Task`**:  Displays a single task (text, completion status, delete button).
- **`TaskList`**:  Renders a list of `Task` components.
- **`TaskForm`**: Handles adding new tasks to the list.

**Step 3: Create the `Task` Component**

- Create a new file: `src/components/Task.js`.
- Implement the `Task` component. Consider how to display the task text and handle the "delete" and "toggle completion" actions:

```jsx
// src/components/Task.js
import React from "react";

// Task component structure - you'll need to fill in the implementation details! 
function Task({ task, onDelete, onToggle }) {
  return ( 
    <div className="task">
      {/* Display the task text */}
      {/* Checkbox for marking tasks as complete */}
      {/* Button to delete the task */}
    </div>
  );
}

export default Task;
```

**Step 4: Create the `TaskList` Component**

- Create a new file: `src/components/TaskList.js`.
- This component will receive an array of tasks and render individual `Task` components:

```jsx
// src/components/TaskList.js
import React from "react";
import Task from "./Task";

// TaskList component - needs implementation
function TaskList({ tasks, onDelete, onToggle }) {
  // How would you map over the 'tasks' array to render Task components? 
  return ( 
    <ul>
       {/* Render Task components here */}
    </ul> 
  );
}

export default TaskList;
```

**Step 5: Create the `TaskForm` Component**

- Create a new file: `src/components/TaskForm.js`.
- This component will manage a form for adding new tasks:

```jsx
// src/components/TaskForm.js
import React, { useState } from "react";

// TaskForm component - needs implementation
function TaskForm({ onAddTask }) {
  // Use useState to manage the input field for the new task

  // Handle form submission - what happens when the user adds a new task? 

  return (
    <form /* onSubmit={...} */ >
      <input 
        type="text" 
        placeholder="Add a new task"
        /* value={...} */
        /* onChange={...} */
      />
      <button type="submit">Add</button>
    </form>
  );
}

export default TaskForm;
```

### Phase 3: Managing State and Logic

**Step 6: Update `App.js` to Manage Tasks**

- In `src/App.js`, you'll manage the application's state (the list of tasks) and pass data and functions down to the child components.

```jsx
// src/App.js
import React, { useState } from "react";
import TaskList from "./components/TaskList";
import TaskForm from "./components/TaskForm";

function App() {
  // Use useState to manage an array of tasks. Each task could be an object with properties like:
  // - id
  // - text 
  // - completed (true/false)

  // Function to add a new task to the state

  // Function to delete a task from the state

  // Function to toggle the 'completed' status of a task

  return (
    <div>
      <h1>My To-Do List</h1> 
      <TaskForm /* onAddTask={...} */ />
      <TaskList /* tasks={...} onDelete={...} onToggle={...} */ />
    </div>
  );
}

export default App;
```

### Phase 4: Styling (Optional)

**Step 7: Add CSS for Basic Styling**

- Add styles to `src/App.css` to customize the look of your app. Consider:
    - Styling the `Task` component (e.g., displaying completed tasks with a line-through).
    - Styling the overall layout of the list.

### Phase 5: Running and Testing the App

**Step 8: Start the Development Server**

- In your terminal, run:

```bash
npm start
```

- This will open the app in your browser. 

**Step 9: Thoroughly Test Your App**
-  Add tasks, mark them as complete, delete them, and make sure your app behaves as expected.

**Important: This guide provides a structured outline. You'll need to fill in the implementation details yourself.  Refer to React documentation, online resources, and tutorials if you get stuck.** 
