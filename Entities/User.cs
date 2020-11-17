using System;

namespace GenerateExcel.Entities
{
    public class User
    {
        public string Id { get; private set; }
        public string Name { get; private set; }
        public string Email { get; private set; }
        public int Age { get; private set; }

        public User(string name, string email, int age)
        {
            Id = Guid.NewGuid().ToString().Substring(0, 8);
            Name = name;
            Email = email;
            Age = age;
        }
    }
}
