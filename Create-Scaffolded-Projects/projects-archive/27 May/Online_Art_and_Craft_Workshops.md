## Online Art and Craft Workshops: React Project Guide

**Project Goal:** Build a web application that connects learners with art instructors for virtual workshops. 

**Project Phases:**

**Phase 1: Basic Setup and Navigation**

**Step 1: Project Setup**

   - Create a new React project using Create React App: `npx create-react-app my-workshop-app`
   - Navigate to the project directory: `cd my-workshop-app`
   - Install necessary dependencies:
     - `npm install @material-ui/core @material-ui/icons react-router-dom react-currency-format`

**Step 2: Basic Structure**

   - Create the following components in the `src` folder:
      - `App.js` (Main app component)
      - `Header.js` (Navigation bar)
      - `HomePage.js` (Homepage with workshop listings)
      - `WorkshopPage.js` (Individual workshop details)
      - `InstructorPage.js` (Instructor profile page)
      - `LoginPage.js` (Login/Registration form)

**Step 3: Routing and Navigation**

   - In `App.js`, use `BrowserRouter`, `Route`, and `Switch` from `react-router-dom` to set up routing.
   - Configure routes for the following pages:
      - `/` (HomePage)
      - `/workshop/:workshopId` (WorkshopPage)
      - `/instructor/:instructorId` (InstructorPage)
      - `/login` (LoginPage)

**Step 4: Header Component**

   - In `Header.js`, design the navigation bar with:
      - A logo or title
      - Links to:
         - Homepage
         - Login/Registration
      - Optional: Search bar for finding workshops

**Step 5: Homepage Component**

   - In `HomePage.js`, display a list of workshops with:
      - Workshop title
      - Instructor name
      - Brief description
      - Image (optional)

**Phase 2: Workshop Details and Instructor Profiles**

**Step 6: Workshop Page Component**

   - In `WorkshopPage.js`, display detailed information about a specific workshop:
      - Title
      - Instructor name and profile picture
      - Full description
      - Dates and times
      - Price
      - Registration button

**Step 7: Instructor Page Component**

   - In `InstructorPage.js`, display an instructor's profile:
      - Name and profile picture
      - Bio
      - List of workshops offered
      - Contact information (optional)

**Phase 3: Registration, Payment, and User Authentication**

**Step 8: Registration Form**

   - In `WorkshopPage.js`, implement a registration form with:
      - User name and email
      - Payment details (using `react-currency-format`)
      - Submit button

**Step 9: Authentication and User Data (Optional)**

   - Choose an authentication provider (Firebase, Auth0, etc.) and integrate it into your app.
   - Store user data in the authentication provider or a backend database.
   - Implement user login and logout functionality.

**Step 10: Payment Integration (Optional)**

   - Integrate a payment gateway (Stripe, PayPal, etc.) to handle secure payment processing.

**Phase 4: Data Fetching and Styling**

**Step 11: Fetching Workshop Data**

   - Create a data source for workshops and instructors (consider using a JSON file, a backend API, or a database).
   - Fetch data from the source and display it in the components using `useEffect` or a data fetching library.

**Step 12: Styling**

   - Use CSS or a CSS framework (Material-UI, Bootstrap, etc.) to style your components and create a visually appealing design.
   - Consider using `styled-components` or `emotion` for more efficient styling.

**Step 13: Responsiveness**

   - Ensure that your app is responsive across different screen sizes using media queries or CSS frameworks.

**Phase 5: Deployment**

**Step 14: Build for Production**

   - Run `npm run build` to create a production-ready build of your application.

**Step 15: Choose a Hosting Platform**

   - Select a hosting provider (Netlify, Vercel, AWS S3, etc.) that supports static websites.

**Step 16: Deploy Your Application**

   - Follow the hosting provider's instructions to deploy your built application.

**Hints and Code Snippets:**

- **Header Component (Header.js)**

```javascript
import React from 'react';
import { Link } from 'react-router-dom';
import AppBar from '@material-ui/core/AppBar';
import Toolbar from '@material-ui/core/Toolbar';
import Typography from '@material-ui/core/Typography';

const Header = () => {
  return (
    <AppBar position="static">
      <Toolbar>
        <Typography variant="h6">
          <Link to="/" style={{ textDecoration: "none", color: "white" }}>
            Online Art & Craft Workshops
          </Link>
        </Typography>
        <Link to="/login" style={{ textDecoration: "none", color: "white", marginLeft: "auto" }}>
          Login
        </Link>
      </Toolbar>
    </AppBar>
  );
};

export default Header;
```

- **Homepage Component (HomePage.js)**

```javascript
import React, { useState, useEffect } from 'react';
import { Link } from 'react-router-dom';
import Grid from '@material-ui/core/Grid';
import Card from '@material-ui/core/Card';
import CardContent from '@material-ui/core/CardContent';
import CardMedia from '@material-ui/core/CardMedia';
import Typography from '@material-ui/core/Typography';

const HomePage = () => {
  const [workshops, setWorkshops] = useState([]);

  useEffect(() => {
    // Fetch workshop data from your source (e.g., JSON file, API)
    fetch('/workshops.json') // Replace with your data source
      .then(res => res.json())
      .then(data => setWorkshops(data));
  }, []);

  return (
    <Grid container spacing={3} style={{ padding: '20px' }}>
      {workshops.map(workshop => (
        <Grid item key={workshop.id} xs={12} sm={6} md={4}>
          <Link to={`/workshop/${workshop.id}`} style={{ textDecoration: "none" }}>
            <Card>
              <CardMedia
                component="img"
                alt={workshop.title}
                height="140"
                image={workshop.image}
                title={workshop.title}
              />
              <CardContent>
                <Typography gutterBottom variant="h5" component="h2">
                  {workshop.title}
                </Typography>
                <Typography variant="body2" color="textSecondary">
                  {workshop.instructor}
                </Typography>
              </CardContent>
            </Card>
          </Link>
        </Grid>
      ))}
    </Grid>
  );
};

export default HomePage;
```

- **Workshop Page Component (WorkshopPage.js)**

```javascript
import React, { useState, useEffect } from 'react';
import { useParams } from 'react-router-dom';
import Typography from '@material-ui/core/Typography';
import Card from '@material-ui/core/Card';
import CardContent from '@material-ui/core/CardContent';
import CardMedia from '@material-ui/core/CardMedia';
import Grid from '@material-ui/core/Grid';
import Button from '@material-ui/core/Button';

const WorkshopPage = () => {
  const { workshopId } = useParams();
  const [workshop, setWorkshop] = useState(null);

  useEffect(() => {
    // Fetch workshop data from your source (e.g., JSON file, API)
    fetch(`/workshops/${workshopId}.json`) // Replace with your data source
      .then(res => res.json())
      .then(data => setWorkshop(data));
  }, [workshopId]);

  if (!workshop) {
    return <div>Loading...</div>;
  }

  return (
    <Grid container spacing={3} style={{ padding: '20px' }}>
      <Grid item xs={12} md={6}>
        <Card>
          <CardMedia
            component="img"
            alt={workshop.title}
            height="400"
            image={workshop.image}
            title={workshop.title}
          />
          <CardContent>
            <Typography variant="h4" gutterBottom>
              {workshop.title}
            </Typography>
            <Typography variant="body2" color="textSecondary">
              {workshop.description}
            </Typography>
            <Typography variant="body1">
              <span style={{ fontWeight: 'bold' }}>Instructor:</span> {workshop.instructor}
            </Typography>
            <Typography variant="body1">
              <span style={{ fontWeight: 'bold' }}>Dates & Times:</span> {workshop.dates}
            </Typography>
            <Typography variant="body1">
              <span style={{ fontWeight: 'bold' }}>Price:</span> ${workshop.price}
            </Typography>
          </CardContent>
        </Card>
      </Grid>
      <Grid item xs={12} md={6}>
        <Typography variant="h5" gutterBottom>
          Register for This Workshop
        </Typography>
        <Button variant="contained" color="primary" onClick={() => {
          // Handle registration logic
          alert('Registering for workshop: ' + workshop.title);
        }}>
          Register
        </Button>
      </Grid>
    </Grid>
  );
};

export default WorkshopPage;
```

- **Instructor Page Component (InstructorPage.js)**

```javascript
import React, { useState, useEffect } from 'react';
import { useParams } from 'react-router-dom';
import Typography from '@material-ui/core/Typography';
import Card from '@material-ui/core/Card';
import CardContent from '@material-ui/core/CardContent';
import CardMedia from '@material-ui/core/CardMedia';
import Grid from '@material-ui/core/Grid';

const InstructorPage = () => {
  const { instructorId } = useParams();
  const [instructor, setInstructor] = useState(null);

  useEffect(() => {
    // Fetch instructor data from your source (e.g., JSON file, API)
    fetch(`/instructors/${instructorId}.json`) // Replace with your data source
      .then(res => res.json())
      .then(data => setInstructor(data));
  }, [instructorId]);

  if (!instructor) {
    return <div>Loading...</div>;
  }

  return (
    <Grid container spacing={3} style={{ padding: '20px' }}>
      <Grid item xs={12} md={6}>
        <Card>
          <CardMedia
            component="img"
            alt={instructor.name}
            height="400"
            image={instructor.image}
            title={instructor.name}
          />
          <CardContent>
            <Typography variant="h4" gutterBottom>
              {instructor.name}
            </Typography>
            <Typography variant="body1" color="textSecondary">
              {instructor.bio}
            </Typography>
            <Typography variant="body1">
              <span style={{ fontWeight: 'bold' }}>Contact:</span> {instructor.email}
            </Typography>
          </CardContent>
        </Card>
      </Grid>
      <Grid item xs={12} md={6}>
        <Typography variant="h5" gutterBottom>
          Workshops Offered
        </Typography>
        <ul>
          {instructor.workshops.map(workshop => (
            <li key={workshop.id}>
              <a href={`/workshop/${workshop.id}`} style={{ textDecoration: "none" }}>{workshop.title}</a>
            </li>
          ))}
        </ul>
      </Grid>
    </Grid>
  );
};

export default InstructorPage;
```
