namespace Goal
{
    public class GoalDetail
    {
        public bool isMainGoal { get; set; }
        public string goalName { get; set; }
        public string goalOwner { get; set; }
        public string target { get; set; }


        public GoalDetail(bool isMainGoal, string goalName, string goalOwner, string target)
        {
            this.isMainGoal = isMainGoal;
            this.goalName = goalName;
            this.goalOwner = goalOwner;
            this.target = target;
        }

    }
}