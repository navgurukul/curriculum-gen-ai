## Interactive Map of Indian Street Art

This project will guide you through building a React application that displays an interactive map of street art locations in India. Users can explore different cities and discover information about the artwork, artists, and their stories.

### Phase 1: Setup and Basic Structure

**Step 1: Project Setup**

* Create a new React project using Create React App:
   ```bash
   npx create-react-app indian-street-art
   cd indian-street-art
   ```

* Install necessary dependencies:
   ```bash
   npm install react-leaflet leaflet
   ```

**Step 2: Project Structure**

* Create the following folders and files:
  ```
  indian-street-art/
  ├── src/
  │   ├── App.js
  │   ├── components/
  │   │   ├── Map.js
  │   │   ├── Marker.js
  │   │   ├── CityInfo.js
  │   │   └── Navigation.js
  │   ├── data/
  │   │   └── street-art-data.js
  │   └── styles/
  │       ├── App.css
  │       ├── Map.css
  │       ├── Marker.css
  │       └── CityInfo.css
  └── public/
      └── index.html
  ```

**Step 3: Basic App Structure**

* In `src/App.js`, define the basic structure of the application:
   ```javascript
   import React from 'react';
   import './styles/App.css';
   import Map from './components/Map';
   import CityInfo from './components/CityInfo';
   import Navigation from './components/Navigation';

   function App() {
     return (
       <div className="App">
         <Navigation />
         <Map />
         <CityInfo />
       </div>
     );
   }

   export default App;
   ```

**Step 4: Basic Styling (Optional)**

* Add basic styling to `src/styles/App.css`:
   ```css
   .App {
     display: flex;
     flex-direction: column;
     height: 100vh;
     background-color: #f2f2f2;
   }
   ```

### Phase 2: Map Component

**Step 1: Map Component Implementation**

* In `src/components/Map.js`, create the map component using Leaflet:
   ```javascript
   import React, { useState, useEffect } from 'react';
   import { MapContainer, TileLayer, Marker, Popup } from 'react-leaflet';
   import './styles/Map.css';
   import streetArtData from '../data/street-art-data';

   const Map = () => {
     const [selectedCity, setSelectedCity] = useState(null);
     const [center, setCenter] = useState([20.5937, 78.9629]);

     useEffect(() => {
       if (selectedCity) {
         const cityData = streetArtData.find((city) => city.name === selectedCity);
         setCenter([cityData.latitude, cityData.longitude]);
       }
     }, [selectedCity]);

     const handleCityChange = (city) => {
       setSelectedCity(city);
     };

     return (
       <MapContainer
         center={center}
         zoom={6}
         style={{ height: '600px', width: '100%', margin: '20px auto' }}
       >
         <TileLayer
           attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
           url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
         />
         {streetArtData.map((city) => (
           <Marker
             key={city.name}
             position={[city.latitude, city.longitude]}
           >
             <Popup>
               <span onClick={() => handleCityChange(city.name)}>
                 {city.name}
               </span>
             </Popup>
           </Marker>
         ))}
       </MapContainer>
     );
   };

   export default Map;
   ```

**Step 2: Data File**

* Create `src/data/street-art-data.js` and populate it with street art data for different cities:
   ```javascript
   const streetArtData = [
     {
       name: 'Mumbai',
       latitude: 19.0760,
       longitude: 72.8777,
       artworks: [
         {
           title: 'Wall Painting',
           artist: 'Artist Name',
           description: 'Description of the artwork',
           imageUrl: 'https://example.com/image1.jpg',
           location: 'Location details'
         },
         // ... more artworks
       ]
     },
     // ... more cities
   ];

   export default streetArtData;
   ```

**Step 3: Basic Map Styling**

* Add basic styling to `src/styles/Map.css`:
   ```css
   .leaflet-container {
     height: 600px;
     width: 100%;
   }
   ```

### Phase 3: Marker Component

**Step 1: Marker Component Implementation**

* In `src/components/Marker.js`, create a marker component that renders individual street art locations:
   ```javascript
   import React from 'react';
   import { Marker, Popup } from 'react-leaflet';
   import './styles/Marker.css';

   const MarkerComponent = ({ artwork, city, handleMarkerClick }) => {
     const { title, artist, imageUrl, location } = artwork;

     return (
       <Marker
         position={[city.latitude, city.longitude]}
         onClick={() => handleMarkerClick(artwork)}
       >
         <Popup>
           <div className="marker-popup">
             <h3>{title}</h3>
             <img src={imageUrl} alt={title} />
             <p>Artist: {artist}</p>
             <p>Location: {location}</p>
           </div>
         </Popup>
       </Marker>
     );
   };

   export default MarkerComponent;
   ```

**Step 2: Marker Component Styling**

* Add styling to `src/styles/Marker.css`:
   ```css
   .marker-popup {
     display: flex;
     flex-direction: column;
     align-items: center;
   }

   .marker-popup img {
     width: 100%;
     max-height: 200px;
     object-fit: cover;
     margin-bottom: 10px;
   }
   ```

### Phase 4: City Info Component

**Step 1: City Info Component Implementation**

* In `src/components/CityInfo.js`, create a component to display information about the selected city and its street art:
   ```javascript
   import React, { useState, useEffect } from 'react';
   import './styles/CityInfo.css';
   import streetArtData from '../data/street-art-data';
   import MarkerComponent from './Marker';

   const CityInfo = () => {
     const [selectedCity, setSelectedCity] = useState(null);
     const [artworks, setArtworks] = useState([]);

     useEffect(() => {
       if (selectedCity) {
         const cityData = streetArtData.find((city) => city.name === selectedCity);
         setArtworks(cityData.artworks);
       }
     }, [selectedCity]);

     const handleMarkerClick = (artwork) => {
       // Handle marker click and update the state
     };

     return (
       <div className="city-info">
         {selectedCity && (
           <div>
             <h2>{selectedCity}</h2>
             <div className="artworks-container">
               {artworks.map((artwork, index) => (
                 <MarkerComponent
                   key={index}
                   artwork={artwork}
                   city={streetArtData.find((city) => city.name === selectedCity)}
                   handleMarkerClick={handleMarkerClick}
                 />
               ))}
             </div>
           </div>
         )}
       </div>
     );
   };

   export default CityInfo;
   ```

**Step 2: City Info Component Styling**

* Add styling to `src/styles/CityInfo.css`:
   ```css
   .city-info {
     padding: 20px;
     flex: 1;
   }

   .artworks-container {
     display: flex;
     flex-wrap: wrap;
     justify-content: center;
   }
   ```

### Phase 5: Navigation Component

**Step 1: Navigation Component Implementation**

* In `src/components/Navigation.js`, create a navigation component to allow users to switch between cities:
   ```javascript
   import React, { useState } from 'react';
   import './styles/Navigation.css';
   import streetArtData from '../data/street-art-data';

   const Navigation = () => {
     const [selectedCity, setSelectedCity] = useState(null);
     const [cities, setCities] = useState([]);

     useEffect(() => {
       setCities(streetArtData.map((city) => city.name));
     }, []);

     const handleCitySelect = (city) => {
       setSelectedCity(city);
     };

     return (
       <nav className="navigation">
         <ul>
           {cities.map((city) => (
             <li key={city} onClick={() => handleCitySelect(city)}>
               {city}
             </li>
           ))}
         </ul>
       </nav>
     );
   };

   export default Navigation;
   ```

**Step 2: Navigation Component Styling**

* Add styling to `src/styles/Navigation.css`:
   ```css
   .navigation {
     padding: 20px;
     background-color: #fff;
     box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
   }

   .navigation ul {
     display: flex;
     list-style: none;
     padding: 0;
   }

   .navigation li {
     margin: 0 10px;
     cursor: pointer;
   }
   ```

### Phase 6: Data Population and Integration

**Step 1: Data Integration**

* In `src/App.js`, pass the selected city state to the `Map` and `CityInfo` components:
   ```javascript
   import React, { useState } from 'react';
   import './styles/App.css';
   import Map from './components/Map';
   import CityInfo from './components/CityInfo';
   import Navigation from './components/Navigation';

   function App() {
     const [selectedCity, setSelectedCity] = useState(null);

     return (
       <div className="App">
         <Navigation setSelectedCity={setSelectedCity} />
         <Map selectedCity={selectedCity} />
         <CityInfo selectedCity={selectedCity} />
       </div>
     );
   }

   export default App;
   ```

**Step 2: Handle City Selection**

* In `src/components/Navigation.js`, handle the city selection and pass the selected city to the parent component:
   ```javascript
   import React, { useState } from 'react';
   import './styles/Navigation.css';
   import streetArtData from '../data/street-art-data';

   const Navigation = ({ setSelectedCity }) => {
     const [cities, setCities] = useState([]);

     useEffect(() => {
       setCities(streetArtData.map((city) => city.name));
     }, []);

     const handleCitySelect = (city) => {
       setSelectedCity(city);
     };

     return (
       <nav className="navigation">
         <ul>
           {cities.map((city) => (
             <li key={city} onClick={() => handleCitySelect(city)}>
               {city}
             </li>
           ))}
         </ul>
       </nav>
     );
   };

   export default Navigation;
   ```

**Step 3: Render Markers Dynamically**

* In `src/components/Map.js`, use the selected city state to render markers for artworks in that city:
   ```javascript
   import React, { useState, useEffect } from 'react';
   import { MapContainer, TileLayer, Marker, Popup } from 'react-leaflet';
   import './styles/Map.css';
   import streetArtData from '../data/street-art-data';
   import MarkerComponent from './Marker';

   const Map = ({ selectedCity }) => {
     const [center, setCenter] = useState([20.5937, 78.9629]);

     useEffect(() => {
       if (selectedCity) {
         const cityData = streetArtData.find((city) => city.name === selectedCity);
         setCenter([cityData.latitude, cityData.longitude]);
       }
     }, [selectedCity]);

     const handleMarkerClick = (artwork) => {
       // Handle marker click and update the state
     };

     return (
       <MapContainer
         center={center}
         zoom={6}
         style={{ height: '600px', width: '100%', margin: '20px auto' }}
       >
         <TileLayer
           attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
           url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
         />
         {selectedCity && (
           <div>
             {streetArtData
               .find((city) => city.name === selectedCity)
               .artworks.map((artwork, index) => (
                 <MarkerComponent
                   key={index}
                   artwork={artwork}
                   city={streetArtData.find((city) => city.name === selectedCity)}
                   handleMarkerClick={handleMarkerClick}
                 />
               ))}
           </div>
         )}
       </MapContainer>
     );
   };

   export default Map;
   ```

**Step 4: Display Artworks in City Info**

* In `src/components/CityInfo.js`, use the selected city state to display information about the city's artworks:
   ```javascript
   import React, { useState, useEffect } from 'react';
   import './styles/CityInfo.css';
   import streetArtData from '../data/street-art-data';
   import MarkerComponent from './Marker';

   const CityInfo = ({ selectedCity }) => {
     const [artworks, setArtworks] = useState([]);

     useEffect(() => {
       if (selectedCity) {
         const cityData = streetArtData.find((city) => city.name === selectedCity);
         setArtworks(cityData.artworks);
       }
     }, [selectedCity]);

     const handleMarkerClick = (artwork) => {
       // Handle marker click and update the state
     };

     return (
       <div className="city-info">
         {selectedCity && (
           <div>
             <h2>{selectedCity}</h2>
             <div className="artworks-container">
               {artworks.map((artwork, index) => (
                 <MarkerComponent
                   key={index}
                   artwork={artwork}
                   city={streetArtData.find((city) => city.name === selectedCity)}
                   handleMarkerClick={handleMarkerClick}
                 />
               ))}
             </div>
           </div>
         )}
       </div>
     );
   };

   export default CityInfo;
   ```

### Phase 7: User Interaction and Functionality

**Step 1: Handle Marker Click**

* In `src/components/Map.js` and `src/components/CityInfo.js`, implement the `handleMarkerClick` function to update the state with selected artwork data:
   ```javascript
   import React, { useState, useEffect } from 'react';
   import { MapContainer, TileLayer, Marker, Popup } from 'react-leaflet';
   import './styles/Map.css';
   import streetArtData from '../data/street-art-data';
   import MarkerComponent from './Marker';

   const Map = ({ selectedCity }) => {
     const [center, setCenter] = useState([20.5937, 78.9629]);
     const [selectedArtwork, setSelectedArtwork] = useState(null);

     useEffect(() => {
       if (selectedCity) {
         const cityData = streetArtData.find((city) => city.name === selectedCity);
         setCenter([cityData.latitude, cityData.longitude]);
       }
     }, [selectedCity]);

     const handleMarkerClick = (artwork) => {
       setSelectedArtwork(artwork);
     };

     return (
       <MapContainer
         center={center}
         zoom={6}
         style={{ height: '600px', width: '100%', margin: '20px auto' }}
       >
         <TileLayer
           attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
           url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
         />
         {selectedCity && (
           <div>
             {streetArtData
               .find((city) => city.name === selectedCity)
               .artworks.map((artwork, index) => (
                 <MarkerComponent
                   key={index}
                   artwork={artwork}
                   city={streetArtData.find((city) => city.name === selectedCity)}
                   handleMarkerClick={handleMarkerClick}
                 />
               ))}
           </div>
         )}
       </MapContainer>
     );
   };

   export default Map;
   ```

   ```javascript
   import React, { useState, useEffect } from 'react';
   import './styles/CityInfo.css';
   import streetArtData from '../data/street-art-data';
   import MarkerComponent from './Marker';

   const CityInfo = ({ selectedCity, selectedArtwork }) => {
     const [artworks, setArtworks] = useState([]);

     useEffect(() => {
       if (selectedCity) {
         const cityData = streetArtData.find((city) => city.name === selectedCity);
         setArtworks(cityData.artworks);
       }
     }, [selectedCity]);

     return (
       <div className="city-info">
         {selectedCity && (
           <div>
             <h2>{selectedCity}</h2>
             {selectedArtwork && (
               <div className="selected-artwork">
                 <h3>{selectedArtwork.title}</h3>
                 <img src={selectedArtwork.imageUrl} alt={selectedArtwork.title} />
                 <p>Artist: {selectedArtwork.artist}</p>
                 <p>Location: {selectedArtwork.location}</p>
               </div>
             )}
             <div className="artworks-container">
               {artworks.map((artwork, index) => (
                 <MarkerComponent
                   key={index}
                   artwork={artwork}
                   city={streetArtData.find((city) => city.name === selectedCity)}
                   handleMarkerClick={handleMarkerClick}
                 />
               ))}
             </div>
           </div>
         )}
       </div>
     );
   };

   export default CityInfo;
   ```

**Step 2: Pass Selected Artwork Data**

* In `src/App.js`, pass the `selectedArtwork` state to the `CityInfo` component:
   ```javascript
   import React, { useState } from 'react';
   import './styles/App.css';
   import Map from './components/Map';
   import CityInfo from './components/CityInfo';
   import Navigation from './components/Navigation';

   function App() {
     const [selectedCity, setSelectedCity] = useState(null);
     const [selectedArtwork, setSelectedArtwork] = useState(null);

     return (
       <div className="App">
         <Navigation setSelectedCity={setSelectedCity} />
         <Map selectedCity={selectedCity} setSelectedArtwork={setSelectedArtwork} />
         <CityInfo selectedCity={selectedCity} selectedArtwork={selectedArtwork} />
       </div>
     );
   }

   export default App;
   ```

**Step 3: Display Selected Artwork**

* In `src/components/CityInfo.js`, display the details of the selected artwork:
   ```javascript
   import React, { useState, useEffect } from 'react';
   import './styles/CityInfo.css';
   import streetArtData from '../data/street-art-data';
   import MarkerComponent from './Marker';

   const CityInfo = ({ selectedCity, selectedArtwork }) => {
     const [artworks, setArtworks] = useState([]);

     useEffect(() => {
       if (selectedCity) {
         const cityData = streetArtData.find((city) => city.name === selectedCity);
         setArtworks(cityData.artworks);
       }
     }, [selectedCity]);

     return (
       <div className="city-info">
         {selectedCity && (
           <div>
             <h2>{selectedCity}</h2>
             {selectedArtwork && (
               <div className="selected-artwork">
                 <h3>{selectedArtwork.title}</h3>
                 <img src={selectedArtwork.imageUrl} alt={selectedArtwork.title} />
                 <p>Artist: {selectedArtwork.artist}</p>
                 <p>Location: {selectedArtwork.location}</p>
               </div>
             )}
             <div className="artworks-container">
               {artworks.map((artwork, index) => (
                 <MarkerComponent
                   key={index}
                   artwork={artwork}
                   city={streetArtData.find((city) => city.name === selectedCity)}
                   handleMarkerClick={handleMarkerClick}
                 />
               ))}
             </div>
           </div>
         )}
       </div>
     );
   };

   export default CityInfo;
   ```

### Phase 8: Enhancements and Additional Features (Optional)

* **Add a search bar:** Allow users to search for specific artworks or artists.
* **Implement a filtering system:** Filter artworks by category, artist, or location.
* **Add user authentication:** Allow users to create accounts and save their favorite artworks.
* **Integrate Google Maps:** Use Google Maps for better map rendering and features.
* **Improve styling and design:** Create a visually appealing and user-friendly interface.
* **Add additional information:** Include more details about artworks, artists, and the street art scene.
* **Use a database:** Store data in a database for persistent storage and easier management.

Remember to populate the `src/data/street-art-data.js` file with actual street art data from India. You can find data through online sources or by directly contacting street artists.

This guide provides a solid foundation for building your interactive map of Indian street art. You can customize it further with additional features and enhancements to create a truly engaging and informative application.