## Building a Community Radio App with React

This guide will walk you through building a community radio app using React. It will focus on essential features like streaming, station selection, and displaying information about the current broadcast.

**Project Phases:**

**Phase 1: Setup & Basic Structure**

**Step 1: Project Setup:**

* Use Create React App to set up a new React project: 
   ```bash
   npx create-react-app community-radio
   cd community-radio 
   ```

* Install necessary dependencies:
   ```bash
   npm install react-router-dom axios 
   ```
   (We'll use axios for API calls)

**Step 2: Project Structure:**

* Create the following files/folders:
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

* Create a basic index.css for overall styles.

**Phase 2: App Shell & Data Fetching**

**Step 3: App.js:**

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

**Step 4: Header.js:**

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

**Step 5: StationList.js:**

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

**Step 6: radioApi.js:**

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

**Phase 3: Player & Station Information**

**Step 7: Player.js:**

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

**Step 8: StationInfo.js:**

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

**Step 9: formatTime.js:**

```javascript
export default function formatTime(seconds) {
  const minutes = Math.floor(seconds / 60);
  const remainingSeconds = Math.floor(seconds % 60);
  return `${minutes.toString().padStart(2, "0")}:${remainingSeconds
    .toString()
    .padStart(2, "0")}`;
}
```

**Phase 4: Styling & Refinement**

**Step 10: Styling:**

* Use CSS to style the components and create a visually appealing user interface.

**Step 11: Additional Features:**

* Implement additional features like:
    * **Station Search:** Allow users to search for stations by name.
    * **Favorites:** Implement a feature to save favorite stations.
    * **More Information:** Display more detailed information about the station or the current broadcast.
    * **User Authentication:** Allow users to create accounts and save preferences.
    * **Social Media Sharing:** Enable users to share their favorite stations on social media.
    * **Accessibility:** Ensure the app is accessible to users with disabilities.

**Remember:**

* Replace placeholders with your API endpoints and actual station data.
* Adapt the code to your specific API structure and requirements.
* Test your app thoroughly and make sure it works as intended on different devices and browsers.
* Consider using a state management library like Redux or Zustand if your app becomes more complex.

This guide provides a solid foundation for building your community radio app. You can expand on these features and customize the design to create a unique and engaging listening experience for your users.