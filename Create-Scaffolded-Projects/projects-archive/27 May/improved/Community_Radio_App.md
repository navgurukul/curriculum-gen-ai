## Building a Community Radio App with React

This guide will walk you through building a community radio app using React. It will focus on essential features like streaming, station selection, and displaying information about the current broadcast. You'll learn to fetch data from a radio station API, manage state, and implement essential React concepts.

**Project Phases:**

**Phase 1: Setting Up Your Project**

**Step 1: Creating Your React Project**

*  Initialize your project:

   ```bash
   npx create-react-app community-radio
   cd community-radio
   ```

*  Install necessary dependencies:

   ```bash
   npm install react-router-dom axios 
   ```

   *  `react-router-dom` allows you to navigate between different parts of your app.
   *  `axios` is a popular library for making API calls.

**Step 2: Project Structure**

*  Create a well-organized file structure for your components:

   ```
   src/
     App.js
     components/
       Header.js
       StationList.js
       Player.js
       StationInfo.js
     styles/
       App.css
       Header.css
       StationList.css
       Player.css
       StationInfo.css
     services/
       radioApi.js 
     utils/
       formatTime.js
   ```

*  Create a basic `index.css` for overall styles.

**Phase 2:  Building the App Shell**

**Step 3: The App.js Component**

```javascript
import React, { useState, useEffect } from "react";
import { BrowserRouter as Router, Routes, Route } from "react-router-dom";
import Header from "./components/Header";
import StationList from "./components/StationList";
import Player from "./components/Player";
import StationInfo from "./components/StationInfo";
import "./styles/App.css";
import radioApi from "./services/radioApi"; 

function App() {
  const [stations, setStations] = useState([]);
  const [currentStation, setCurrentStation] = useState(null);
  const [isPlaying, setIsPlaying] = useState(false);
  const [currentTime, setCurrentTime] = useState(0); 

  useEffect(() => {
    const fetchStations = async () => {
      try {
        const data = await radioApi.getStations();
        setStations(data);
      } catch (error) {
        console.error("Error fetching stations:", error);
      }
    };
    fetchStations();
  }, []);

  const handleStationSelect = (station) => {
    setCurrentStation(station);
    setIsPlaying(true); 
  };

  const handlePlayPause = () => {
    setIsPlaying(!isPlaying); 
  };

  const handleTimeUpdate = (time) => {
    setCurrentTime(time);
  };

  return (
    <Router>
      <div className="App">
        <Header />

        <Routes>
          <Route path="/" element={<StationList stations={stations} onSelect={handleStationSelect} />} />
          <Route path="/station/:id" element={<StationInfo station={currentStation} isPlaying={isPlaying} currentTime={currentTime} />} />
        </Routes>

        {currentStation && (
          <Player
            station={currentStation}
            isPlaying={isPlaying}
            onPlayPause={handlePlayPause}
            onTimeUpdate={handleTimeUpdate}
          />
        )}
      </div>
    </Router>
  );
}

export default App;
```

**Explanation:**

*   The `App` component serves as the main container for your app.
*   It uses state variables to store:
    *   `stations`: An array of station objects
    *   `currentStation`: The currently selected station
    *   `isPlaying`:  Indicates if the audio is playing
    *   `currentTime`: The current playback time of the audio
*   The `useEffect` hook fetches the list of stations from the API when the component mounts.
*   The `handleStationSelect` function updates the `currentStation` and starts playback.
*   The `handlePlayPause` function toggles the playback state.
*   The `handleTimeUpdate` function updates the `currentTime` state.

**Step 4: The Header Component (Header.js)**

```javascript
import React from "react";
import { Link } from "react-router-dom";
import "./styles/Header.css";

function Header() {
  return (
    <header className="header">
      <Link to="/" className="header__title">
        <h1>Community Radio</h1>
      </Link>
    </header>
  );
}

export default Header;
```

*   The `Header` component displays a simple title that links to the home page.

**Step 5: Displaying Station Information (StationList.js)**

```javascript
import React from "react";
import "./styles/StationList.css";

function StationList({ stations, onSelect }) {
  return (
    <ul className="station-list">
      {stations.map((station) => (
        <li key={station.id} onClick={() => onSelect(station)}>
          {station.name}
        </li>
      ))}
    </ul>
  );
}

export default StationList;
```

*   The `StationList` component renders a list of radio stations.
*   Each station is represented as a list item (`<li>`) and is clickable to trigger the `onSelect` function.

**Step 6:  Fetching Station Data (radioApi.js)**

```javascript
import axios from "axios";

const API_BASE_URL = "https://your-radio-api-base-url.com"; // Replace with your API endpoint

const radioApi = {
  getStations: async () => {
    const response = await axios.get(`${API_BASE_URL}/stations`);
    return response.data;
  },
  // ... (Add other API methods if needed)
};

export default radioApi;
```

*   The `radioApi` module provides functions for interacting with your radio station API.
*   **Important:** Replace `"https://your-radio-api-base-url.com"` with the actual base URL of your API.

**Phase 3: Implementing the Player & Station Details**

**Step 7: The Audio Player Component (Player.js)**

```javascript
import React, { useRef, useEffect } from "react";
import "./styles/Player.css";

function Player({ station, isPlaying, onPlayPause, onTimeUpdate }) {
  const audioRef = useRef(null);

  useEffect(() => {
    if (audioRef.current) {
      if (isPlaying) {
        audioRef.current.play();
      } else {
        audioRef.current.pause();
      }
    }
  }, [isPlaying]);

  useEffect(() => {
    if (audioRef.current) {
      audioRef.current.ontimeupdate = () => {
        onTimeUpdate(audioRef.current.currentTime);
      };
    }
  }, []);

  return (
    <div className="player">
      <audio ref={audioRef} src={station.streamUrl} />
      <button onClick={onPlayPause}>
        {isPlaying ? "Pause" : "Play"}
      </button>
    </div>
  );
}

export default Player;
```

*   The `Player` component handles audio playback.
*   It uses a `useRef` hook to store a reference to the `<audio>` element.
*   It updates the audio playback state based on the `isPlaying` prop.
*   It listens for time updates from the audio element using `ontimeupdate`.

**Step 8: Displaying Station Information (StationInfo.js)**

```javascript
import React from "react";
import { useParams } from "react-router-dom";
import formatTime from "../utils/formatTime";
import "./styles/StationInfo.css";

function StationInfo({ station, isPlaying, currentTime }) {
  const { id } = useParams();

  if (!station) {
    return <div>Loading...</div>;
  }

  return (
    <div className="station-info">
      <h2>{station.name}</h2>
      <p>{station.description}</p>

      {isPlaying && (
        <div>
          <span>Now Playing: </span>
          <span>{station.currentBroadcast}</span>
        </div>
      )}

      <div>
        {isPlaying ? (
          <span>Time Elapsed: {formatTime(currentTime)}</span>
        ) : (
          <span>Stream not playing</span>
        )}
      </div>
    </div>
  );
}

export default StationInfo;
```

*   The `StationInfo` component displays information about the currently selected station.
*   It uses the `useParams` hook to access the station ID from the URL.
*   It shows the station name, description, and optionally the current broadcast details.
*   It uses the `formatTime` utility to format the current playback time.

**Step 9:  Formatting Time (formatTime.js)**

```javascript
export default function formatTime(seconds) {
  const minutes = Math.floor(seconds / 60);
  const remainingSeconds = Math.floor(seconds % 60);
  return `${minutes.toString().padStart(2, "0")}:${remainingSeconds
    .toString()
    .padStart(2, "0")}`;
}
```

*   The `formatTime` utility converts seconds into a human-readable time format (minutes:seconds).

**Phase 4: Styling & Refinement**

**Step 10: Styling Your Components**

*   Use CSS within the `styles/` directory to style your components: `App.css`, `Header.css`, `StationList.css`, `Player.css`, and `StationInfo.css`.
*   Make sure your app has a visually appealing and user-friendly interface.

**Step 11: Adding Additional Features (Optional)**

*   **Station Search:** Implement a search bar to allow users to search for stations by name.
*   **Favorites:** Enable users to save their favorite stations.
*   **More Information:** Show more detailed station information or broadcast details.
*   **User Authentication:** Implement user accounts to save preferences.
*   **Social Media Sharing:** Allow users to share stations on social media.
*   **Accessibility:** Ensure your app is accessible for users with disabilities.

**Important Reminders:**

*   Replace the placeholder API URL with the correct endpoint for your radio station API.
*   Adapt your code to work with the specific structure and data provided by your API.
*   Test your app thoroughly in different browsers and devices to ensure it works correctly.
*   If your app becomes more complex, consider using a state management library like Redux or Zustand.

**Next Steps**

This guide has given you a strong foundation for building your community radio app. By implementing the steps above, you'll have a working app that fetches data, plays audio, and displays station information. Now it's your turn to expand on these features and customize the design to create a unique and enjoyable listening experience for your users. Good luck!
