// theme.js - Manage theme switching and local persistence
(function () {
  // Determine initial theme
  const savedTheme = localStorage.getItem('theme') || 'light';
  document.documentElement.setAttribute('data-theme', savedTheme);

  // Set class on document body or html element for theme indicators
  document.addEventListener('DOMContentLoaded', () => {
    updateToggleIcons(savedTheme);
  });
})();

// Toggle theme function
function toggleTheme() {
  const currentTheme = document.documentElement.getAttribute('data-theme') || 'light';
  const newTheme = currentTheme === 'light' ? 'dark' : 'light';
  
  document.documentElement.setAttribute('data-theme', newTheme);
  localStorage.setItem('theme', newTheme);
  
  // Update UI icons across the page
  updateToggleIcons(newTheme);
  
  // Dispatch custom event for page-specific listeners (e.g. Chart.js redraws)
  const event = new CustomEvent('theme-changed', { detail: { theme: newTheme } });
  window.dispatchEvent(event);
}

// Update the theme toggle button icons if they exist
function updateToggleIcons(theme) {
  const lightIcons = document.querySelectorAll('.theme-icon-light');
  const darkIcons = document.querySelectorAll('.theme-icon-dark');
  
  if (theme === 'dark') {
    lightIcons.forEach(el => el.style.display = 'inline');
    darkIcons.forEach(el => el.style.display = 'none');
  } else {
    lightIcons.forEach(el => el.style.display = 'none');
    darkIcons.forEach(el => el.style.display = 'inline');
  }
}
