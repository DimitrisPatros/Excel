namespace Excel
{
    public class Student
    {
        public string Name { get; set; }
        public string Course { get; set; }
        public int RegisterId { get; set; }

        public Student(string name, string course, int registerId)
        {
            Name = name;
            Course = course;
            RegisterId = registerId;
        }
  
        public override string ToString()
        {
            return $"Name: {Name}\t  Course: {Course}\t  Register ID: {RegisterId}";
        }
    }
}
