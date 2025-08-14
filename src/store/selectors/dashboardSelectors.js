export const selectHoveredCard = (state) => state.dashboard?.hoveredCard || null;
export const selectHoveredButton = (state) => state.dashboard?.hoveredButton || null;
export const selectHoveredLink = (state) => state.dashboard?.hoveredLink || null;
export const selectIsMobile = (state) => state.dashboard?.isMobile || false;
export const selectCurrentTime = (state) => state.dashboard?.currentTime || '';
export const selectSearchTerm = (state) => state.dashboard?.searchTerm || '';
export const selectSelectedCategory = (state) => state.dashboard?.selectedCategory || 'all';
export const selectNotifications = (state) => state.dashboard?.notifications || [];
export const selectShowNotifications = (state) => state.dashboard?.showNotifications || false;
export const selectStats = (state) => state.dashboard?.stats || { usersOnline: 0, tasksCompleted: 0 };
export const selectIsDarkMode = (state) => state.dashboard?.isDarkMode || false;
export const selectShowQuickTips = (state) => state.dashboard?.showQuickTips || true;
export const selectCurrentTip = (state) => state.dashboard?.currentTip || 0;
export const selectShowSearch = (state) => state.dashboard?.showSearch || false;
export const selectRecentActivity = (state) => state.dashboard?.recentActivity || [];

export const selectFilteredTools = (state, tools) => {
  const searchTerm = selectSearchTerm(state);
  const selectedCategory = selectSelectedCategory(state);
  
  if (!tools) return [];
  
  return tools.filter(tool => {
    const matchesSearch = !searchTerm || 
      tool.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
      tool.description.toLowerCase().includes(searchTerm.toLowerCase()) ||
      tool.tag.toLowerCase().includes(searchTerm.toLowerCase());
    
    const matchesCategory = selectedCategory === 'all' || tool.category === selectedCategory;
    
    return matchesSearch && matchesCategory;
  });
};