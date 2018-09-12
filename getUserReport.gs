function getUserReport(report, user, userIndex) {
  switch (report) {
    case 'created_tasks':
      return getСreatedTasks(user);
      break;

    case 'feedback_tasks':
      return getFeedbackTasks(user);
      break;

    case 'boss_rating_avg':
      return getBossRatingAverage(user);
      break;
  }
}

function getСreatedTasks(user) {
  var res = APIRequest('issues', {query: [
    {key: 'tracker_id', value: '!5'},
    {key: 'author_id', value: user.id},
    {key: 'status_id', value: '*'},
    {key: 'created_on', value: getDateRage(OPTIONS.startDate, OPTIONS.finalDate)}
  ]});
  return res.issues;
}

function getFeedbackTasks(user) {
  var allTask = [];

  for (var i = 4; i <= 5; i++) {
    var withFeedback = APIRequest('issues', {query: [
      {key: 'tracker_id', value: '!5'},
      {key: 'author_id', value: user.id},
      {key: 'status_id', value: i},
      {key: 'cf_34', value: '1'},
      {key: 'created_on', value: getDateRage(OPTIONS.startDate, OPTIONS.finalDate)}
    ]});

    allTask = allTask.concat(withFeedback.issues);
  }

  for (var i = 3; i <= 5; i++) {
    var noFeedback = APIRequest('issues', {query: [
      {key: 'tracker_id', value: '!5'},
      {key: 'author_id', value: user.id},
      {key: 'status_id', value: i},
      {key: 'cf_35', value: '1'},
      {key: 'created_on', value: getDateRage(OPTIONS.startDate, OPTIONS.finalDate)}
    ]});

    allTask = allTask.concat(noFeedback.issues);
  }

  return allTask;
}

function getBossRatingAverage(user) {
  var res = APIRequest('issues', {query: [
    {key: 'tracker_id', value: '!5'},
    {key: 'author_id', value: user.id},
    {key: 'status_id', value: '*'},
    {key: 'created_on', value: getDateRage(OPTIONS.startDate, OPTIONS.finalDate)},
    {key: 'cf_8', value: '*'}
  ]});

  var sum = res.issues.reduce(function(a, c) {
    return a + parseInt(c.custom_fields.find(function(i) {return i.id === 8}).value, 10);
  }, 0);

  return res.issues.length ? sum / res.issues.length : 0;
}
