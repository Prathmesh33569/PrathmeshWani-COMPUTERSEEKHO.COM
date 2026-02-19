/**
 * Check if followup date is pending (before today)
 */
export const isPending = (followupDate) => {
  return new Date(followupDate) < new Date(new Date().setHours(0, 0, 0, 0));
};

/**
 * Check if followup date is today
 */
export const isToday = (followupDate) => {
  const today = new Date().toISOString().split('T')[0];
  return followupDate === today;
};

/**
 * Get default next followup date (+3 days)
 */
export const getDefaultNextDate = () => {
  const date = new Date();
  date.setDate(date.getDate() + 3);
  return date.toISOString().split('T')[0];
};

/**
 * Format date to Indian locale
 */
export const formatDate = (dateString) => {
  return new Date(dateString).toLocaleDateString('en-IN');
};

/**
 * Get status class based on followup date
 */
export const getStatusClass = (followupDate) => {
  if (isPending(followupDate)) {
    return 'border-red-500 bg-red-50';
  } else if (isToday(followupDate)) {
    return 'border-green-500';
  }
  return 'border-blue-500';
};

/**
 * Get status text color based on followup date
 */
export const getStatusTextClass = (followupDate) => {
  if (isPending(followupDate)) {
    return 'text-red-600';
  } else if (isToday(followupDate)) {
    return 'text-green-600';
  }
  return 'text-blue-600';
};

/**
 * Get status label
 */
export const getStatusLabel = (followupDate) => {
  if (isPending(followupDate)) {
    return ' (Overdue)';
  } else if (isToday(followupDate)) {
    return ' (Today)';
  }
  return '';
};