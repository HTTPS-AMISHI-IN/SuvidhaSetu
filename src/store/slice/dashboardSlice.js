// store/slices/dashboardSlice.js
// ================================
import { createSlice } from '@reduxjs/toolkit';

const dashboardSlice = createSlice({
  name: 'dashboard',
  initialState: {
    hoveredCard: null,
    hoveredButton: null,
    hoveredLink: null,
    isMobile: false,
    currentTime: '',
    searchTerm: '',
    selectedCategory: 'all',
    showNotifications: false,
    showSearch: false,
    showQuickTips: true,
    currentTip: 0,
    isDarkMode: false,
    stats: {
      usersOnline: 42,
      tasksCompleted: 128
    },
    notifications: [
      {
        type: 'info',
        title: 'Welcome to Dashboard',
        message: 'All systems are operational'
      }
    ],
    recentActivity: []
  },
  reducers: {
    setHoveredCard: (state, action) => {
      state.hoveredCard = action.payload;
    },
    setHoveredButton: (state, action) => {
      state.hoveredButton = action.payload;
    },
    setHoveredLink: (state, action) => {
      state.hoveredLink = action.payload;
    },
    setIsMobile: (state, action) => {
      state.isMobile = action.payload;
    },
    setCurrentTime: (state, action) => {
      state.currentTime = action.payload;
    },
    setSearchTerm: (state, action) => {
      state.searchTerm = action.payload;
    },
    setSelectedCategory: (state, action) => {
      state.selectedCategory = action.payload;
    },
    toggleNotifications: (state) => {
      state.showNotifications = !state.showNotifications;
    },
    setShowSearch: (state, action) => {
      state.showSearch = action.payload;
    },
    toggleSearch: (state) => {
      state.showSearch = !state.showSearch;
    },
    toggleQuickTips: (state) => {
      state.showQuickTips = !state.showQuickTips;
    },
    setCurrentTip: (state, action) => {
      state.currentTip = action.payload;
    },
    nextTip: (state) => {
      state.currentTip = (state.currentTip + 1) % 6; // Assuming 6 tips
    },
    toggleDarkMode: (state) => {
      state.isDarkMode = !state.isDarkMode;
    },
    updateStats: (state, action) => {
      state.stats = { ...state.stats, ...action.payload };
    },
    addNotification: (state, action) => {
      state.notifications.push(action.payload);
    },
    removeNotification: (state, action) => {
      state.notifications = state.notifications.filter((_, index) => index !== action.payload);
    },
    addRecentActivity: (state, action) => {
      state.recentActivity.unshift(action.payload);
      if (state.recentActivity.length > 10) {
        state.recentActivity = state.recentActivity.slice(0, 10);
      }
    }
  }
});

export const {
  setHoveredCard,
  setHoveredButton,
  setHoveredLink,
  setIsMobile,
  setCurrentTime,
  setSearchTerm,
  setSelectedCategory,
  toggleNotifications,
  setShowSearch,
  toggleSearch,
  toggleQuickTips,
  setCurrentTip,
  nextTip,
  toggleDarkMode,
  updateStats,
  addNotification,
  removeNotification,
  addRecentActivity
} = dashboardSlice.actions;

export default dashboardSlice.reducer;