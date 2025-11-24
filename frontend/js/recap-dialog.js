<script>
  // Recap Dialog JavaScript
  let formData = {};
  
  // Initialize the form
  window.onload = function() {
    setupFormValidation();
    loadFormData();
  };
  
  // Setup form validation and event handlers
  function setupFormValidation() {
    const form = document.getElementById('recapForm');
    const requiredFields = form.querySelectorAll('[required]');
    
    // Add real-time validation
    requiredFields.forEach(field => {
      field.addEventListener('blur', validateField);
      field.addEventListener('input', clearFieldError);
    });
    
    // Handle form submission
    form.addEventListener('submit', handleFormSubmit);
  }
  
  // Load existing form data
  function loadFormData() {
    // Auto-populate any additional data if needed
    const focusSkill = document.getElementById('focusSkill');
    const wins = document.getElementById('wins');
    
    // If focus skill is empty, try to use first line of wins
    if (focusSkill && wins && !focusSkill.value && wins.value) {
      const firstLine = wins.value.split('\n')[0].replace(/^[•\-\*]\s*/, '').trim();
      if (firstLine.length < 80) { // Reasonable focus skill length
        focusSkill.value = firstLine;
      }
    }
  }
  
  // Validate individual field
  function validateField(event) {
    const field = event.target;
    const value = field.value.trim();
    
    // Remove existing error styling
    field.classList.remove('error');
    
    // Check required fields
    if (field.hasAttribute('required') && !value) {
      showFieldError(field, 'This field is required');
      return false;
    }
    
    // Validate email format
    if (field.type === 'email' && value) {
      if (!isValidEmail(value)) {
        showFieldError(field, 'Please enter a valid email address');
        return false;
      }
    }
    
    // Validate multiple emails
    if (field.id === 'additionalEmails' && value) {
      const emails = value.split(',').map(email => email.trim());
      for (let email of emails) {
        if (email && !isValidEmail(email)) {
          showFieldError(field, `Invalid email: ${email}`);
          return false;
        }
      }
    }
    
    return true;
  }
  
  // Clear field error styling
  function clearFieldError(event) {
    const field = event.target;
    field.classList.remove('error');
    
    // Remove error message if exists
    const errorMsg = field.parentNode.querySelector('.field-error');
    if (errorMsg) {
      errorMsg.remove();
    }
  }
  
  // Show field error
  function showFieldError(field, message) {
    field.classList.add('error');
    
    // Remove existing error message
    const existingError = field.parentNode.querySelector('.field-error');
    if (existingError) {
      existingError.remove();
    }
    
    // Add error message
    const errorDiv = document.createElement('div');
    errorDiv.className = 'field-error';
    errorDiv.style.color = '#d93025';
    errorDiv.style.fontSize = '12px';
    errorDiv.style.marginTop = '4px';
    errorDiv.textContent = message;
    
    field.parentNode.appendChild(errorDiv);
  }
  
  // Handle form submission
  function handleFormSubmit(event) {
    event.preventDefault();
    
    // Validate all fields
    const form = document.getElementById('recapForm');
    const requiredFields = form.querySelectorAll('[required]');
    let isValid = true;
    
    requiredFields.forEach(field => {
      if (!validateField({ target: field })) {
        isValid = false;
      }
    });
    
    if (!isValid) {
      showStatus('Please fix the errors above', 'error');
      return;
    }
    
    // Collect form data
    formData = {
      parentEmail: document.getElementById('parentEmail').value.trim(),
      additionalEmails: document.getElementById('additionalEmails').value.trim(),
      focusSkill: document.getElementById('focusSkill').value.trim(),
      wins: document.getElementById('wins').value.trim(),
      skillsMastered: document.getElementById('skillsMastered').value.trim(),
      skillsPracticed: document.getElementById('skillsPracticed').value.trim(),
      skillsIntroduced: document.getElementById('skillsIntroduced').value.trim(),
      struggles: document.getElementById('struggles').value.trim(),
      homework: document.getElementById('homework').value.trim(),
      nextSession: document.getElementById('nextSession').value.trim()
    };
    
    // Show loading and send email
    showLoading(true);
    
    google.script.run
      .withSuccessHandler(handleSendSuccess)
      .withFailureHandler(handleSendError)
      .sendSessionRecapEmail(formData);
  }
  
  // Handle successful email send
  function handleSendSuccess(result) {
    showLoading(false);
    showStatus('Session recap sent successfully!', 'success');
    
    // Auto-close dialog after 2 seconds
    setTimeout(() => {
      google.script.host.close();
    }, 2000);
  }
  
  // Handle send error
  function handleSendError(error) {
    showLoading(false);
    showStatus(error.message || 'Failed to send recap. Please try again.', 'error');
  }
  
  // Preview email
  function previewEmail() {
    // Collect current form data
    const currentData = {
      parentEmail: document.getElementById('parentEmail').value.trim(),
      additionalEmails: document.getElementById('additionalEmails').value.trim(),
      focusSkill: document.getElementById('focusSkill').value.trim(),
      wins: document.getElementById('wins').value.trim(),
      skillsMastered: document.getElementById('skillsMastered').value.trim(),
      skillsPracticed: document.getElementById('skillsPracticed').value.trim(),
      skillsIntroduced: document.getElementById('skillsIntroduced').value.trim(),
      struggles: document.getElementById('struggles').value.trim(),
      homework: document.getElementById('homework').value.trim(),
      nextSession: document.getElementById('nextSession').value.trim()
    };
    
    // Generate preview
    google.script.run
      .withSuccessHandler(() => {
        // Preview dialog opened
      })
      .withFailureHandler((error) => {
        showStatus('Failed to generate preview', 'error');
      })
      .showEmailPreview(currentData);
  }
  
  // Save as draft
  function saveRecap() {
    const currentData = {
      focusSkill: document.getElementById('focusSkill').value.trim(),
      wins: document.getElementById('wins').value.trim(),
      skillsMastered: document.getElementById('skillsMastered').value.trim(),
      skillsPracticed: document.getElementById('skillsPracticed').value.trim(),
      skillsIntroduced: document.getElementById('skillsIntroduced').value.trim(),
      struggles: document.getElementById('struggles').value.trim(),
      homework: document.getElementById('homework').value.trim(),
      nextSession: document.getElementById('nextSession').value.trim()
    };
    
    google.script.run
      .withSuccessHandler(() => {
        showStatus('Draft saved successfully', 'success');
      })
      .withFailureHandler(() => {
        showStatus('Failed to save draft', 'error');
      })
      .saveQuickNotes(currentData);
  }
  
  // Open Acuity Scheduling
  function openAcuity() {
    window.open('https://secure.acuityscheduling.com/admin/clients', '_blank');
  }
  
  // Show status message
  function showStatus(message, type) {
    const statusDiv = document.getElementById('statusMessage');
    statusDiv.className = `status-message ${type}`;
    statusDiv.textContent = message;
    statusDiv.style.display = 'block';
    
    // Auto-hide after 5 seconds for non-error messages
    if (type !== 'error') {
      setTimeout(() => {
        statusDiv.style.display = 'none';
      }, 5000);
    }
  }
  
  // Show/hide loading spinner
  function showLoading(show) {
    const loadingDiv = document.getElementById('loadingSpinner');
    const buttonGroup = document.querySelector('.button-group');
    
    if (show) {
      loadingDiv.classList.add('show');
      buttonGroup.querySelectorAll('button').forEach(btn => {
        btn.disabled = true;
      });
    } else {
      loadingDiv.classList.remove('show');
      buttonGroup.querySelectorAll('button').forEach(btn => {
        btn.disabled = false;
      });
    }
  }
  
  // Email validation
  function isValidEmail(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  }
  
  // Auto-save functionality
  let autoSaveTimeout;
  function setupAutoSave() {
    const inputs = document.querySelectorAll('input, textarea');
    inputs.forEach(input => {
      input.addEventListener('input', () => {
        clearTimeout(autoSaveTimeout);
        autoSaveTimeout = setTimeout(saveRecap, 30000); // Auto-save after 30 seconds
      });
    });
  }
  
  // Initialize auto-save
  setupAutoSave();
</script>