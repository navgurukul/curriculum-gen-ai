## Building a To-Do List App with React

This guide will walk you through building a simple To-Do List app using React.

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
- Replace the contents of `src/App.js` with:

```jsx
import React from "react";

function App() {
  return (
    <div>
      {/* Your app content will go here */}
    </div>
  );
}

export default App;

```

### Phase 2: Building the Components

**Step 3: Create the `Task` Component**

- Create a new file `src/components/Task.js`.
- This component represents a single task in the list:

```jsx
import React from "react";

function Task({ task, onDelete, onToggle }) {
  return (
    <div className="task">
      <input
        type="checkbox"
        checked={task.completed}
        onChange={() => onToggle(task.id)}
      />
      <span className={task.completed ? "completed" : ""}>{task.text}</span>
      <button onClick={() => onDelete(task.id)}>Delete</button>
    </div>
  );
}

export default Task;

```

**Step 4: Create the `TaskList` Component**

- Create a new file `src/components/TaskList.js`.
- This component displays the list of tasks:

```jsx
import React from "react";
import Task from "./Task";

function TaskList({ tasks, onDelete, onToggle }) {
  return (
    <ul>
      {tasks.map((task) => (
        <li key={task.id}>
          <Task task={task} onDelete={onDelete} onToggle={onToggle} />
        </li>
      ))}
    </ul>
  );
}

export default TaskList;

```

**Step 5: Create the `TaskForm` Component**

- Create a new file `src/components/TaskForm.js`.
- This component handles adding new tasks:

```jsx
import React, { useState } from "react";

function TaskForm({ onAddTask }) {
  const [newTask, setNewTask] = useState("");

  const handleSubmit = (e) => {
    e.preventDefault();
    if (newTask.trim() !== "") {
      onAddTask(newTask);
      setNewTask("");
    }
  };

  return (
    <form onSubmit={handleSubmit}>
      <input
        type="text"
        placeholder="Add a new task"
        value={newTask}
        onChange={(e) => setNewTask(e.target.value)}
      />
      <button type="submit">Add</button>
    </form>
  );
}

export default TaskForm;

```

### Phase 3: Managing State and Logic

**Step 6: Update `App.js` to Manage Tasks**

- Import necessary components and set up state to store the tasks:

```jsx
import React, { useState } from "react";
import TaskList from "./components/TaskList";
import TaskForm from "./components/TaskForm";

function App() {
  const [tasks, setTasks] = useState([]);

  // Function to add a new task
  const addTask = (newTask) => {
    setTasks([
      ...tasks,
      {
        id: Date.now(),
        text: newTask,
        completed: false,
      },
    ]);
  };

  // Function to delete a task
  const deleteTask = (taskId) => {
    setTasks(tasks.filter((task) => task.id !== taskId));
  };

  // Function to toggle task completion
  const toggleTask = (taskId) => {
    setTasks(
      tasks.map((task) =>
        task.id === taskId ? { ...task, completed: !task.completed } : task
      )
    );
  };

  return (
    <div>
      <h1>To-Do List</h1>
      <TaskForm onAddTask={addTask} />
      <TaskList tasks={tasks} onDelete={deleteTask} onToggle={toggleTask} />
    </div>
  );
}

export default App;

```

### Phase 4: Styling (Optional)

**Step 7: Add CSS for Basic Styling**

- Add styles to `src/App.css` to customize the look of your app:

```css
.task {
  display: flex;
  align-items: center;
  margin-bottom: 10px;
}

.task span {
  margin-left: 10px;
}

.task.completed span {
  text-decoration: line-through;
  color: gray;
}
```

### Phase 5: Running the App

**Step 8: Start the Development Server**

- In your terminal, run:

```bash
npm start
```

- This will start the development server and open the app in your browser. Now you can start adding, completing, and deleting tasks!

This guide provides a basic structure for a To-Do List app. You can further enhance it by implementing features like local storage persistence, task filtering, and more advanced styling.
