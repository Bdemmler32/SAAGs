document.addEventListener('DOMContentLoaded', function() {
  // DOM Elements
  const dayButtonsContainer = document.getElementById('dayButtons');
  const eventTypeButtonsContainer = document.getElementById('eventTypeButtons');
  const scheduleGrid = document.getElementById('scheduleGrid');
  const noEventsMessage = document.getElementById('noEvents');
  const dateInfo = document.getElementById('date-info');
  const loadingIndicator = document.getElementById('loading');
  const errorMessage = document.getElementById('errorMessage');
  const expandCollapseToggle = document.getElementById('expandCollapseToggle');
  const exportPdfBtn = document.getElementById('exportPdfBtn');
  
  // Search-related DOM elements
  const searchIcon = document.getElementById('searchIcon');
  const searchContainer = document.getElementById('searchContainer');
  const searchInput = document.getElementById('searchInput');
  const clearSearchBtn = document.getElementById('clearSearchBtn');
  const searchResultsInfo = document.getElementById('searchResultsInfo');
  
  // State variables
  let events = [];
  let filteredEvents = [];
  let selectedDay = null;
  let selectedEventTypes = new Set(); // Multi-select for event types
  let lastUpdated = '';
  let isSearchActive = false;
  let prevFilterState = null;
  
  // Event Type Colors - updated with new distinct palette
  const eventTypeColors = {
    'All Conference Activities': { bg: '#f5f0e8', border: '#d4c4a8' }, // Light Brown
    'Council/Committee Meetings': { bg: '#fff2e6', border: '#ffccb3' }, // Light Orange
    'Networking and Social Functions': { bg: '#fff9e6', border: '#fff0b3' }, // Light Yellow
    'Networking ans Social Functions': { bg: '#fff9e6', border: '#fff0b3' }, // Handle typo - Light Yellow
    'Other (Workshop/Course etcâ€¦)': { bg: '#e6fff2', border: '#b3ffd6' }, // Light Green
    'Registration': { bg: '#f0e6ff', border: '#d6b3ff' }, // Light Purple
    'Technical Program': { bg: '#ffe6e6', border: '#ffb3b3' }, // Light Red
    'Ticketed Event': { bg: '#e6f4ff', border: '#b3d7ff' }, // Light Blue
    'Session': { bg: '#f8f9fa', border: '#e9ecef' } // Light Gray for sessions
  };
  
  // Initialize
  fetchScheduleData();
  
  // Fetch schedule data from Excel file
  async function fetchScheduleData() {
    // Show loading indicator
    loadingIndicator.style.display = 'block';
    noEventsMessage.style.display = 'none';
    scheduleGrid.style.display = 'none';
    errorMessage.style.display = 'none';
    
    try {
      // Fetch the Excel file
      const response = await fetch('MiNES2025SAAG.xlsx');
      if (!response.ok) {
        throw new Error('Network response was not ok');
      }
      const excelData = await response.arrayBuffer();
      
      // Parse the Excel file
      const workbook = XLSX.read(new Uint8Array(excelData), {
        cellDates: true
      });
      
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      // Get the update date from D1
      const updateDateCell = worksheet['D1'];
      if (updateDateCell && updateDateCell.v) {
        const updateDate = new Date(updateDateCell.v);
        lastUpdated = updateDate.toLocaleDateString();
      }
      
      // Get the data starting from row 3 (header row)
      events = XLSX.utils.sheet_to_json(worksheet, {
        range: 2 // Start from row 3 (index 2)
      });
      
      // Format time fields and process event types + nested sessions
      events = events.map(event => {
        const eventTypes = event["Event Type"] ? event["Event Type"].split(';').map(t => t.trim()) : [];
        const primaryEventType = eventTypes.find(t => t !== 'Ticketed Event') || eventTypes[0];
        const isTicketed = eventTypes.includes('Ticketed Event');
        const isSession = primaryEventType === 'Session';
        
        return {
          ...event,
          "Event": event["Event/Function"], // Map the new column name to the expected field
          "Time Start": formatExcelTime(event["Time Start"]),
          "Time End": formatExcelTime(event["Time End"]),
          "EventTypes": eventTypes,
          "PrimaryEventType": primaryEventType,
          "IsTicketed": isTicketed,
          "IsSession": isSession,
          "PDFLink": event["PDF Link"] || null
        };
      });
      
      // Process nested sessions - group sessions under preceding parent events
      const processedEvents = [];
      let currentParent = null;
      
      events.forEach(event => {
        if (event.IsSession) {
          // This is a session - add it to the current parent's sessions
          if (currentParent) {
            if (!currentParent.sessions) {
              currentParent.sessions = [];
            }
            currentParent.sessions.push(event);
          }
        } else {
          // This is a regular event - could be a parent for future sessions
          currentParent = event;
          processedEvents.push(event);
        }
      });
      
      // Update events to the processed list (without standalone sessions)
      events = processedEvents;
      
      // Sort events chronologically by date and time
      events.sort((a, b) => {
        // First compare dates
        const dateA = new Date(a.Date.split(',')[1] + ',' + a.Date.split(',')[0]);
        const dateB = new Date(b.Date.split(',')[1] + ',' + b.Date.split(',')[0]);
        
        if (dateA.getTime() !== dateB.getTime()) {
          return dateA - dateB;
        }
        
        // If same date, sort by time
        return timeToMinutes(a["Time Start"]) - timeToMinutes(b["Time Start"]);
      });
      
      filteredEvents = [...events];
      
      // Update date info
      dateInfo.textContent = `Current as of ${lastUpdated}`;
      
      // Hide loading indicator
      loadingIndicator.style.display = 'none';
      scheduleGrid.style.display = 'flex';
      
      // Initialize UI
      createDayButtons();
      createEventTypeButtons();
      renderSchedule();
      setupSearchFunctionality();
    } catch (error) {
      // Show error message
      loadingIndicator.style.display = 'none';
      errorMessage.style.display = 'block';
      console.error('Error fetching schedule data:', error);
    }
  }
  
  // Format Excel date/time to 12-hour format (1:00 PM)
  function formatExcelTime(excelTime) {
    if (!excelTime) return '';
    
    // Check if it's already a string (properly formatted)
    if (typeof excelTime === 'string' && !excelTime.includes('T')) {
      return excelTime;
    }
    
    let date;
    if (typeof excelTime === 'string' && excelTime.includes('T')) {
      // ISO string format
      date = new Date(excelTime);
    } else if (excelTime instanceof Date) {
      date = excelTime;
    } else {
      return excelTime; // Return as is if we can't handle it
    }
    
    // Format to 12-hour time
    let hours = date.getHours();
    const minutes = date.getMinutes();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    
    hours = hours % 12;
    hours = hours ? hours : 12; // 0 should be 12
    const minutesStr = minutes < 10 ? '0' + minutes : minutes;
    
    return `${hours}:${minutesStr} ${ampm}`;
  }
  
  // Create day filter buttons
  function createDayButtons() {
    // Clear container
    dayButtonsContainer.innerHTML = '';
    
    // All days button
    const allDaysBtn = document.createElement('button');
    allDaysBtn.type = 'button';
    allDaysBtn.className = 'sched-day-button active';
    allDaysBtn.textContent = 'All Days';
    allDaysBtn.addEventListener('click', function() {
      setActiveDay(null, this);
    });
    dayButtonsContainer.appendChild(allDaysBtn);
    
    // Get all unique days from the events
    const uniqueDays = [...new Set(events.map(event => event.Date))];
    
    // Sort uniqueDays chronologically
    uniqueDays.sort((a, b) => {
      const dateA = new Date(a.split(',')[1] + ',' + a.split(',')[0]);
      const dateB = new Date(b.split(',')[1] + ',' + b.split(',')[0]);
      return dateA - dateB;
    });
    
    // Day-specific buttons
    uniqueDays.forEach(day => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'sched-day-button';
      btn.textContent = day.split(',')[0]; // Just the day name
      btn.addEventListener('click', function() {
        setActiveDay(day, this);
      });
      dayButtonsContainer.appendChild(btn);
    });
  }
  
  // Create event type filter buttons
  function createEventTypeButtons() {
    // Clear container
    eventTypeButtonsContainer.innerHTML = '';
    
    // Get all unique event types from the events (excluding Session type)
    const allEventTypes = new Set();
    events.forEach(event => {
      event.EventTypes.forEach(type => {
        if (type && type !== 'Session') allEventTypes.add(type);
      });
    });
    
    const uniqueEventTypes = Array.from(allEventTypes).sort();
    
    // Create buttons for each event type
    uniqueEventTypes.forEach(eventType => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'sched-event-type-button';
      btn.textContent = eventType;
      btn.dataset.eventType = eventType;
      
      // Set button color based on event type
      const colors = eventTypeColors[eventType] || eventTypeColors['Technical Program'];
      btn.style.backgroundColor = colors.bg;
      btn.style.borderColor = colors.border;
      btn.style.border = `2px solid ${colors.border}`;
      
      btn.addEventListener('click', function() {
        toggleEventTypeFilter(eventType, this);
      });
      
      // Add active state styling
      btn.addEventListener('mousedown', function() {
        if (!selectedEventTypes.has(eventType)) {
          this.style.backgroundColor = colors.border;
        }
      });
      
      btn.addEventListener('mouseup', function() {
        if (!selectedEventTypes.has(eventType)) {
          this.style.backgroundColor = colors.bg;
        }
      });
      
      eventTypeButtonsContainer.appendChild(btn);
    });
  }
  
  // Toggle event type filter (multi-select)
  function toggleEventTypeFilter(eventType, button) {
    const colors = eventTypeColors[eventType] || eventTypeColors['Technical Program'];
    
    if (selectedEventTypes.has(eventType)) {
      // Remove from selection
      selectedEventTypes.delete(eventType);
      button.classList.remove('active');
      button.style.backgroundColor = colors.bg; // Reset to original background
    } else {
      // Add to selection
      selectedEventTypes.add(eventType);
      button.classList.add('active');
      button.style.backgroundColor = colors.border; // Set to border color when active
    }
    
    if (isSearchActive) {
      // If search is active, apply both filters and search
      performSearch(searchInput.value);
    } else {
      // Otherwise just apply filters
      applyFilters();
    }
  }
  
  // Set active day
  function setActiveDay(day, button) {
    selectedDay = day;
    
    // Update button styles
    const buttons = dayButtonsContainer.querySelectorAll('.sched-day-button');
    buttons.forEach(btn => btn.classList.remove('active'));
    button.classList.add('active');
    
    if (isSearchActive) {
      // If search is active, apply both day and search filters
      performSearch(searchInput.value);
    } else {
      // Otherwise just apply day filter
      applyFilters();
    }
  }
  
  // Apply day and event type filters
  function applyFilters() {
    let filtered = [...events];
    
    // Apply day filter
    if (selectedDay) {
      filtered = filtered.filter(event => event.Date === selectedDay);
    }
    
    // Apply event type filter (multi-select)
    if (selectedEventTypes.size > 0) {
      filtered = filtered.filter(event => {
        return event.EventTypes.some(type => selectedEventTypes.has(type));
      });
    }
    
    filteredEvents = filtered;
    renderSchedule();
  }
  
  // Set up search functionality
  function setupSearchFunctionality() {
    // Toggle search container visibility
    searchIcon.addEventListener('click', toggleSearch);
    
    // Close search when X button is clicked
    clearSearchBtn.addEventListener('click', clearSearch);
    
    // Perform search as user types
    searchInput.addEventListener('input', function() {
      performSearch(this.value);
    });
    
    // Focus input when search is opened
    searchInput.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') {
        clearSearch();
      }
    });
  }
  
  // Toggle search container visibility
  function toggleSearch() {
    if (!isSearchActive) {
      // Opening search - save current state
      prevFilterState = {
        selectedDay: selectedDay,
        selectedEventTypes: new Set(selectedEventTypes),
        filteredEvents: [...filteredEvents]
      };
      
      // Show search container with animation
      searchContainer.style.display = 'block';
      searchInput.focus();
      isSearchActive = true;
      
      // Hide search icon when search is active
      searchIcon.style.display = 'none';
    } else {
      clearSearch();
    }
  }
  
  // Clear search and close search container
  function clearSearch() {
    // Clear input
    searchInput.value = '';
    searchResultsInfo.textContent = '';
    
    // Hide search container
    searchContainer.style.display = 'none';
    isSearchActive = false;
    
    // Show search icon again
    searchIcon.style.display = 'block';
    
    // Restore previous state
    if (prevFilterState) {
      selectedDay = prevFilterState.selectedDay;
      selectedEventTypes = new Set(prevFilterState.selectedEventTypes);
      
      // Update day button selection
      const dayButtons = dayButtonsContainer.querySelectorAll('.sched-day-button');
      dayButtons.forEach(btn => btn.classList.remove('active'));
      
      // Find the correct day button to activate
      if (selectedDay === null) {
        // Activate "All Days" button
        dayButtons[0].classList.add('active');
      } else {
        // Find and activate the correct day button
        const dayButtonText = selectedDay.split(',')[0];
        for (let i = 1; i < dayButtons.length; i++) {
          if (dayButtons[i].textContent === dayButtonText) {
            dayButtons[i].classList.add('active');
            break;
          }
        }
      }
      
      // Update event type button selection
      const eventTypeButtons = eventTypeButtonsContainer.querySelectorAll('.sched-event-type-button');
      eventTypeButtons.forEach(btn => {
        const eventType = btn.dataset.eventType;
        const colors = eventTypeColors[eventType] || eventTypeColors['Technical Program'];
        
        if (selectedEventTypes.has(eventType)) {
          btn.classList.add('active');
          btn.style.backgroundColor = colors.border; // Set to border color when active
        } else {
          btn.classList.remove('active');
          btn.style.backgroundColor = colors.bg; // Reset to original background
        }
      });
      
      // Apply the restored filter state
      applyFilters();
    }
  }
  
  // Perform search on events
  function performSearch(query) {
    if (!query.trim()) {
      // If search query is empty, just apply day filter
      if (isSearchActive) {
        searchResultsInfo.textContent = '';
        applyFilters();
      }
      return;
    }
    
    // Start with all events or filtered events
    let baseEvents = [...events];
    if (selectedDay) {
      baseEvents = baseEvents.filter(event => event.Date === selectedDay);
    }
    
    // Apply event type filter
    if (selectedEventTypes.size > 0) {
      baseEvents = baseEvents.filter(event => {
        return event.EventTypes.some(type => selectedEventTypes.has(type));
      });
    }
    
    // Convert query to lowercase for case-insensitive search
    const searchTerms = query.toLowerCase().trim();
    
    // Filter events based on search terms
    const results = baseEvents.filter(event => {
      // Search in event title
      const titleMatch = event.Event && event.Event.toString().toLowerCase().includes(searchTerms);
      
      // Search in event details
      const detailsMatch = event["Event Details"] && 
        event["Event Details"].toString().toLowerCase().includes(searchTerms);
      
      // Search in location
      const locationMatch = event.Location && 
        event.Location.toString().toLowerCase().includes(searchTerms);
      
      // Search in event type
      const typeMatch = event["Event Type"] && 
        event["Event Type"].toString().toLowerCase().includes(searchTerms);
      
      // Search in nested sessions
      let sessionMatch = false;
      if (event.sessions && event.sessions.length > 0) {
        sessionMatch = event.sessions.some(session => {
          return (session.Event && session.Event.toString().toLowerCase().includes(searchTerms)) ||
                 (session["Event Details"] && session["Event Details"].toString().toLowerCase().includes(searchTerms)) ||
                 (session.Location && session.Location.toString().toLowerCase().includes(searchTerms));
        });
      }
      
      // Return true if any field matches
      return titleMatch || detailsMatch || locationMatch || typeMatch || sessionMatch;
    });
    
    // Update results info
    if (isSearchActive) {
      searchResultsInfo.textContent = `Showing ${results.length} result${results.length !== 1 ? 's' : ''} for "${query}"`;
    }
    
    // Update filtered events and render
    filteredEvents = results;
    renderSchedule();
    
    // Show message if no results
    if (results.length === 0) {
      noEventsMessage.style.display = 'block';
    } else {
      noEventsMessage.style.display = 'none';
    }
  }
  
  // Render the schedule grid
  function renderSchedule() {
    scheduleGrid.innerHTML = '';
    
    if (filteredEvents.length === 0) {
      noEventsMessage.style.display = 'block';
      return;
    } else {
      noEventsMessage.style.display = 'none';
    }
    
    // Group events by day
    const eventsByDay = {};
    
    // Get all unique days from filtered events
    const uniqueDays = [...new Set(filteredEvents.map(event => event.Date))].sort((a, b) => {
      const dateA = new Date(a.split(',')[1] + ',' + a.split(',')[0]);
      const dateB = new Date(b.split(',')[1] + ',' + b.split(',')[0]);
      return dateA - dateB;
    });
    
    // Create groups of events by day
    uniqueDays.forEach(day => {
      eventsByDay[day] = filteredEvents.filter(event => event.Date === day);
    });
    
    // Create sections for days with events
    uniqueDays.forEach(day => {
      const dayEvents = eventsByDay[day];
      
      if (dayEvents.length === 0) return;
      
      // Create day section
      const section = document.createElement('div');
      section.className = 'sched-day-section';
      
      // Create header
      const header = document.createElement('div');
      header.className = 'sched-day-header';
      header.textContent = day;
      section.appendChild(header);
      
      // Create content container
      const content = document.createElement('div');
      content.className = 'sched-day-content';
      
      // Sort events by time
      dayEvents.sort((a, b) => {
        return timeToMinutes(a["Time Start"]) - timeToMinutes(b["Time Start"]);
      });
      
      // Add events
      dayEvents.forEach(event => {
        const eventEl = createEventElement(event);
        content.appendChild(eventEl);
      });
      
      section.appendChild(content);
      scheduleGrid.appendChild(section);
    });
  }
  
  // Create an event element
  function createEventElement(event) {
    const primaryEventType = event.PrimaryEventType;
    const isTicketed = event.IsTicketed;
    const isNetworking = primaryEventType === "Networking and Social Functions" || primaryEventType === "Networking ans Social Functions";
    const isSetup = primaryEventType === "Setup";
    
    // Get colors for this event type
    const colors = eventTypeColors[primaryEventType] || eventTypeColors['Technical Program'];
    
    const element = document.createElement('div');
    element.className = `sched-event`;
    element.style.backgroundColor = colors.bg;
    element.style.borderColor = colors.border;
    
    if (isTicketed) {
      element.classList.add('ticketed');
    }
    if (isNetworking || isSetup) {
      element.classList.add('sched-italic-title');
    }
    
    // Event title - this will always be shown
    const title = document.createElement('div');
    title.className = 'sched-event-title';
    title.innerHTML = `<span>${event.Event}</span>`;
    
    // Time display - this will always be shown
    const time = document.createElement('div');
    time.className = 'sched-event-time';
    time.textContent = `${event["Time Start"]} - ${event["Time End"]}`;
    
    element.appendChild(title);
    element.appendChild(time);
    
    // Create event details container - initially hidden, shown when expanded
    // Only add details if Event Details has content OR if event has sessions
    const hasEventDetails = event["Event Details"] && event["Event Details"].toString().trim() !== '';
    const hasSessions = event.sessions && event.sessions.length > 0;
    
    if (hasEventDetails || hasSessions) {
      const details = document.createElement('div');
      details.className = 'sched-event-details';
      details.style.borderTopColor = colors.border; // Match border color
      
      let detailsHTML = '';
      
      // Add regular event details if they exist
      if (hasEventDetails) {
        const eventDetails = event["Event Details"].replace(/\n/g, '<br>');
        detailsHTML += `
          <div><strong>Event Details:</strong> ${eventDetails}</div>
          <div><strong>Location:</strong> ${event.Location || 'TBD'}</div>
          <div><strong>Event Type:</strong> ${event["Event Type"]}</div>
        `;
      }
      
      // Add nested sessions if they exist
      if (hasSessions) {
        detailsHTML += `
          <div class="nested-sessions">
            <div class="nested-sessions-header">
              <div class="session-icon"></div>
              Session Presentations (${event.sessions.length})
            </div>
        `;
        
        event.sessions.forEach(session => {
          const sessionDetails = session["Event Details"] ? session["Event Details"].replace(/\n/g, '<br>') : '';
          detailsHTML += `
            <div class="nested-session">
              <div class="session-title">${session.Event}</div>
              ${session["Time Start"] ? `<div class="session-time">${session["Time Start"]} - ${session["Time End"]}</div>` : ''}
              ${sessionDetails ? `<div class="session-description">${sessionDetails}</div>` : ''}
              ${session.Location ? `<div class="session-location"><strong>Location:</strong> ${session.Location}</div>` : ''}
              ${session.PDFLink ? `<a href="${session.PDFLink}" class="session-link" target="_blank" onclick="event.stopPropagation();">📄 View PDF</a>` : ''}
            </div>
          `;
        });
        
        detailsHTML += `</div>`;
      }
      
      details.innerHTML = detailsHTML;
      element.appendChild(details);
    }
    
    // Add ticketed badge if applicable
    if (isTicketed) {
      const badge = document.createElement('div');
      badge.className = 'sched-ticketed-badge';
      
      const badgeText = document.createElement('div');
      badgeText.className = 'sched-ticketed-text';
      badgeText.textContent = 'TICKETED';
      
      badge.appendChild(badgeText);
      element.appendChild(badge);
    }
    
    // Add click event to expand/collapse for events that have details or sessions
    if (hasEventDetails || hasSessions) {
      element.addEventListener('click', function(e) {
        // Toggle expanded state
        this.classList.toggle('expanded');
      });
    }
    
    return element;
  }
  
  // Expand/Collapse toggle affects all events
  expandCollapseToggle.addEventListener('change', function() {
    const allEvents = document.querySelectorAll('.sched-event');
    
    if (this.checked) {
      // Expand all events
      allEvents.forEach(event => {
        event.classList.add('expanded');
      });
    } else {
      // Collapse all events
      allEvents.forEach(event => {
        event.classList.remove('expanded');
      });
    }
  });
  
  // PDF Export functionality
  exportPdfBtn.addEventListener('click', function() {
    try {
      // Create a clone of the schedule to modify for PDF export
      const originalContainer = document.getElementById('schedule-container');
      const pdfContainer = originalContainer.cloneNode(true);
      
      // Generate filename with current date
      let pdfFilename = 'MiNES2025-SAAG';
      
      // Extract the date from the Current as of text
      const dateText = dateInfo.textContent;
      if (dateText && dateText.includes('Current as of')) {
        const dateMatch = dateText.match(/Current as of ([\d\/]+)/);
        if (dateMatch && dateMatch[1]) {
          // Convert date format to MMDDYYYY
          const dateParts = dateMatch[1].split('/');
          if (dateParts.length === 3) {
            const month = dateParts[0].padStart(2, '0');
            const day = dateParts[1].padStart(2, '0');
            const year = dateParts[2];
            pdfFilename += `-${month}${day}${year}`;
          }
        }
      }
      
      // Add .pdf extension
      pdfFilename += '.pdf';
      
      // Set the container to a fixed width for PDF export
      pdfContainer.style.width = '1100px';
      pdfContainer.style.maxWidth = 'none';
      pdfContainer.style.margin = '0';
      pdfContainer.style.padding = '10px';
      pdfContainer.style.backgroundColor = 'white';
      
      // Remove filters container (keep only the header)
      const filtersContainer = pdfContainer.querySelector('.sched-filters-container');
      if (filtersContainer) {
        filtersContainer.innerHTML = '';
        filtersContainer.style.display = 'none';
      }
      
      // Show header image for PDF export and ensure it's full width
      const headerImage = pdfContainer.querySelector('.sched-header-image');
      if (headerImage) {
        headerImage.style.display = 'block';
        headerImage.style.width = '100%';
        headerImage.style.textAlign = 'center';
        headerImage.style.marginBottom = '15px';
        headerImage.style.borderRadius = '8px';
        headerImage.style.overflow = 'hidden';
        headerImage.style.boxShadow = '0 2px 5px rgba(0,0,0,0.15)';
        headerImage.style.border = '1px solid rgba(0,0,0,0.1)';
        
        // Use the PDF-specific header image
        const img = headerImage.querySelector('img');
        if (img) {
          img.src = 'MiNES2025Webheader-ExportPDF.jpg';
          img.style.width = '100%';
          img.style.maxWidth = '1050px';
          img.style.height = 'auto';
          img.style.margin = '0 auto';
          img.style.display = 'block';
          img.style.borderRadius = '8px';
        }
      }
      
      // Get all unique days from the events
      const uniqueDays = [...new Set(events.map(event => event.Date))];
      
      // Sort uniqueDays chronologically
      uniqueDays.sort((a, b) => {
        const dateA = new Date(a.split(',')[1] + ',' + a.split(',')[0]);
        const dateB = new Date(b.split(',')[1] + ',' + b.split(',')[0]);
        return dateA - dateB;
      });
      
      // Determine number of columns based on number of days
      const numDays = uniqueDays.length;
      
      // Clear current schedule grid and set to align at top
      const scheduleGridPdf = pdfContainer.querySelector('#scheduleGrid');
      if (scheduleGridPdf) {
        scheduleGridPdf.innerHTML = '';
        scheduleGridPdf.style.display = 'grid';
        scheduleGridPdf.style.gridTemplateColumns = `repeat(${numDays}, 1fr)`;
        scheduleGridPdf.style.gap = '5px';
        scheduleGridPdf.style.marginTop = '10px';
        scheduleGridPdf.style.alignItems = 'start';
      }
      
      // Check if the events are currently expanded or collapsed
      const areEventsExpanded = expandCollapseToggle.checked;
      
      // Group events by day
      const eventsByDay = {};
      uniqueDays.forEach(day => {
        eventsByDay[day] = events.filter(event => event.Date === day);
      });
      
      // Create a column for each day
      uniqueDays.forEach(day => {
        const dayEvents = eventsByDay[day];
        
        // Create day column
        const dayColumn = document.createElement('div');
        dayColumn.className = 'day-column';
        dayColumn.style.width = '100%';
        dayColumn.style.overflow = 'hidden';
        dayColumn.style.display = 'flex';
        dayColumn.style.flexDirection = 'column';
        dayColumn.style.alignItems = 'stretch';
        
        // Create day header
        const dayHeader = document.createElement('div');
        dayHeader.className = 'day-header';
        
        // Split the date parts
        const dateParts = day.split(',');
        const dayName = dateParts[0].trim();
        const dateDetail = dateParts[1].trim();
        
        // Create header content with day and date on same line with comma
        dayHeader.innerHTML = `${dayName}, ${dateDetail}`;
        
        dayHeader.style.backgroundColor = '#333';
        dayHeader.style.color = 'white';
        dayHeader.style.padding = '8px 4px';
        dayHeader.style.textAlign = 'center';
        dayHeader.style.borderRadius = '5px 5px 0 0';
        dayHeader.style.fontWeight = 'bold';
        dayHeader.style.fontSize = '11px';
        dayHeader.style.height = '32px';
        dayHeader.style.display = 'flex';
        dayHeader.style.alignItems = 'center';
        dayHeader.style.justifyContent = 'center';
        
        dayColumn.appendChild(dayHeader);
        
        // Create events container
        const eventsContainer = document.createElement('div');
        eventsContainer.className = 'day-events';
        eventsContainer.style.backgroundColor = 'white';
        eventsContainer.style.padding = '4px';
        eventsContainer.style.borderRadius = '0 0 5px 5px';
        eventsContainer.style.flexGrow = '1';
        eventsContainer.style.display = 'flex';
        eventsContainer.style.flexDirection = 'column';
        eventsContainer.style.gap = '3px';
        
        // Sort events by time
        dayEvents.sort((a, b) => {
          return timeToMinutes(a["Time Start"]) - timeToMinutes(b["Time Start"]);
        });
        
        // Add events to container
        dayEvents.forEach(event => {
          const primaryEventType = event.PrimaryEventType;
          const isTicketed = event.IsTicketed;
          const isNetworking = primaryEventType === "Networking and Social Functions" || primaryEventType === "Networking ans Social Functions";
          const isSetup = primaryEventType === "Setup";
          
          // Get colors for this event type
          const colors = eventTypeColors[primaryEventType] || eventTypeColors['Technical Program'];
          
          const eventEl = document.createElement('div');
          eventEl.className = `event-pdf`;
          if (areEventsExpanded) {
            eventEl.classList.add('expanded');
          }
          
          eventEl.style.padding = '4px';
          eventEl.style.borderRadius = '3px';
          eventEl.style.fontSize = '8px';
          eventEl.style.position = 'relative';
          eventEl.style.marginBottom = '2px';
          eventEl.style.lineHeight = '1.2';
          eventEl.style.backgroundColor = colors.bg;
          eventEl.style.border = `1px solid ${colors.border}`;
          
          // Add ticketed indicator if needed
          if (isTicketed) {
            const indicator = document.createElement('div');
            indicator.style.position = 'absolute';
            indicator.style.right = '0';
            indicator.style.top = '0';
            indicator.style.bottom = '0';
            indicator.style.width = '6px';
            indicator.style.backgroundColor = '#4a7aff';
            indicator.style.borderTopRightRadius = '3px';
            indicator.style.borderBottomRightRadius = '3px';
            
            eventEl.style.position = 'relative';
            eventEl.style.paddingRight = '8px';
            
            eventEl.appendChild(indicator);
          }
          
          // Event title
          const titleEl = document.createElement('div');
          titleEl.style.fontWeight = 'bold';
          titleEl.style.marginBottom = '2px';
          
          // Apply italic style for Networking and Setup events
          if (isNetworking || isSetup) {
            titleEl.style.fontStyle = 'italic';
          }
          
          titleEl.textContent = event.Event;
          
          // Event time
          const timeEl = document.createElement('div');
          timeEl.style.fontSize = '8px';
          timeEl.style.color = '#444';
          timeEl.textContent = `${event["Time Start"]} - ${event["Time End"]}`;
          
          eventEl.appendChild(titleEl);
          eventEl.appendChild(timeEl);
          
          // If expanded and has details or sessions, add details
          const hasEventDetails = event["Event Details"] && event["Event Details"].toString().trim() !== '';
          const hasSessions = event.sessions && event.sessions.length > 0;
          
          if (areEventsExpanded && (hasEventDetails || hasSessions)) {
            const detailsEl = document.createElement('div');
            detailsEl.style.marginTop = '4px';
            detailsEl.style.borderTop = `1px solid ${colors.border}`;
            detailsEl.style.paddingTop = '4px';
            detailsEl.style.fontSize = '7px';
            
            let detailsHTML = '';
            
            // Add regular event details if they exist
            if (hasEventDetails) {
              const eventDetailsForPdf = event["Event Details"].replace(/\n/g, '<br>');
              detailsHTML += `
                <div><strong>Event Details:</strong> ${eventDetailsForPdf}</div>
                <div><strong>Location:</strong> ${event.Location || 'TBD'}</div>
                <div><strong>Event Type:</strong> ${event["Event Type"]}</div>
              `;
            }
            
            // Add sessions for PDF if they exist
            if (hasSessions) {
              detailsHTML += `<div><strong>Sessions:</strong></div>`;
              event.sessions.forEach(session => {
                detailsHTML += `<div style="margin-left: 8px;">• ${session.Event}`;
                if (session["Time Start"]) {
                  detailsHTML += ` (${session["Time Start"]} - ${session["Time End"]})`;
                }
                detailsHTML += `</div>`;
              });
            }
            
            detailsEl.innerHTML = detailsHTML;
            eventEl.appendChild(detailsEl);
          }
          
          eventsContainer.appendChild(eventEl);
        });
        
        dayColumn.appendChild(eventsContainer);
        if (scheduleGridPdf) {
          scheduleGridPdf.appendChild(dayColumn);
        }
      });
      
      // Add explanatory legend at the bottom - now shows event type colors
      const legendRow = document.createElement('div');
      legendRow.style.display = 'flex';
      legendRow.style.justifyContent = 'center';
      legendRow.style.gap = '10px';
      legendRow.style.marginTop = '8px';
      legendRow.style.fontSize = '7px';
      legendRow.style.flexWrap = 'wrap';
      
      // Get unique event types from events for legend
      const uniqueEventTypes = [...new Set(events.map(e => e.PrimaryEventType))].sort();
      
      uniqueEventTypes.forEach(eventType => {
        const colors = eventTypeColors[eventType] || eventTypeColors['Technical Program'];
        const legendItem = document.createElement('div');
        legendItem.style.display = 'flex';
        legendItem.style.alignItems = 'center';
        legendItem.style.marginBottom = '2px';
        
        const colorBox = document.createElement('span');
        colorBox.style.width = '8px';
        colorBox.style.height = '8px';
        colorBox.style.backgroundColor = colors.bg;
        colorBox.style.border = `1px solid ${colors.border}`;
        colorBox.style.display = 'inline-block';
        colorBox.style.marginRight = '3px';
        
        legendItem.appendChild(colorBox);
        legendItem.appendChild(document.createTextNode(eventType.length > 20 ? eventType.substring(0, 20) + '...' : eventType));
        
        legendRow.appendChild(legendItem);
      });
      
      // Ticketed legend
      const ticketedLegend = document.createElement('div');
      ticketedLegend.style.display = 'flex';
      ticketedLegend.style.alignItems = 'center';
      ticketedLegend.style.marginBottom = '2px';
      const ticketedColor = document.createElement('span');
      ticketedColor.style.width = '8px';
      ticketedColor.style.height = '8px';
      ticketedColor.style.border = '1px solid #ccc';
      ticketedColor.style.borderRightWidth = '4px';
      ticketedColor.style.borderRightColor = '#4a7aff';
      ticketedColor.style.display = 'inline-block';
      ticketedColor.style.marginRight = '3px';
      ticketedLegend.appendChild(ticketedColor);
      ticketedLegend.appendChild(document.createTextNode('Ticketed Event'));
      
      legendRow.appendChild(ticketedLegend);
      
      // Add legend row to container
      pdfContainer.appendChild(legendRow);
      
      // Temporarily add the cloned container to the document for rendering
      pdfContainer.style.position = 'absolute';
      pdfContainer.style.left = '-9999px';
      document.body.appendChild(pdfContainer);
      
      // Use html2canvas to capture the container
      html2canvas(pdfContainer, {
        scale: 2.5,
        useCORS: true,
        logging: false,
        width: 1100,
        imageTimeout: 0,
        backgroundColor: '#ffffff',
        letterRendering: true,
        allowTaint: true,
        useCORS: true
      }).then(function(canvas) {
        try {
          // Remove the temporary container
          document.body.removeChild(pdfContainer);
          
          // Create PDF in landscape orientation (11x8.5 inches)
          const { jsPDF } = window.jspdf;
          const pdf = new jsPDF({
            orientation: 'landscape',
            unit: 'in',
            format: 'letter',
            compress: true
          });
          
          // Calculate the scaling ratio to fit the canvas to the PDF
          const imgWidth = 11 - 0.4;
          const imgHeight = 8.5 - 0.4;
          const canvasRatio = canvas.height / canvas.width;
          
          let finalWidth = imgWidth;
          let finalHeight = imgWidth * canvasRatio;
          
          // Adjust if the image is too tall
          if (finalHeight > imgHeight) {
            finalHeight = imgHeight;
            finalWidth = imgHeight / canvasRatio;
          }
          
          // Position at the top of the page instead of centering vertically
          const offsetX = (11 - finalWidth) / 2;
          const offsetY = 0.2;
          
          // Add the image to the PDF with quality settings
          const imgData = canvas.toDataURL('image/png', 1.0);
          pdf.addImage(imgData, 'PNG', offsetX, offsetY, finalWidth, finalHeight, undefined, 'FAST');
          
          // Save the PDF
          pdf.save(pdfFilename);
          
        } catch (innerError) {
          console.error("Error in PDF generation:", innerError);
          alert("Error creating PDF: " + innerError.message);
          if (document.body.contains(pdfContainer)) {
            document.body.removeChild(pdfContainer);
          }
        }
      }).catch(function(canvasError) {
        console.error("Error in html2canvas:", canvasError);
        alert("Error capturing page: " + canvasError.message);
        if (document.body.contains(pdfContainer)) {
          document.body.removeChild(pdfContainer);
        }
      });
    } catch (outerError) {
      console.error("Error in PDF export:", outerError);
      alert("Error starting PDF export: " + outerError.message);
    }
  });
  
  // Convert time to minutes for sorting
  function timeToMinutes(timeStr) {
    if (!timeStr) return 0;
    
    const timeParts = timeStr.split(' ');
    const hourMin = timeParts[0].split(':');
    const hour = parseInt(hourMin[0]);
    const minutes = parseInt(hourMin[1]);
    const isPM = timeParts[1] === 'PM';
    
    let hour24 = hour;
    if (isPM && hour !== 12) hour24 = hour + 12;
    if (!isPM && hour === 12) hour24 = 0;
    
    return hour24 * 60 + minutes;
  }
});