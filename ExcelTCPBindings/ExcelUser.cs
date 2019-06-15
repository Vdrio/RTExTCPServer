using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace ExcelTCPBindings
{
    public enum UserAccessLevel
    {
        Default = 0, 
        Admin = 1
    }

    [Serializable]
    public class ExcelUser:ISerializable
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
        public UserAccessLevel AccessLevel
        {
            get { return accessLevel; }
            set {
                if (accessLevel != value)
                {
                    AccessLevelChangedEventArgs e = new AccessLevelChangedEventArgs { OldValue = accessLevel, NewValue = value };
                    OnAccessLevelChanged?.Invoke(this, e);
                }
                accessLevel = value;
            }
        }
        UserAccessLevel accessLevel;

        public Guid Id { get; set; }

        public event EventHandler OnAccessLevelChanged;

        public ExcelUser()
        {
            Id = Guid.NewGuid();
            //OnAccessLevelChanged += AccessLevelChanged;
        }

        public ExcelUser(string firstName, string lastName, UserAccessLevel accessLevel = UserAccessLevel.Default)
        {
            Id = Guid.NewGuid();
            FirstName = firstName;
            LastName = lastName;
            AccessLevel = accessLevel;
            //OnAccessLevelChanged += AccessLevelChanged;
        }

        protected ExcelUser(SerializationInfo info, StreamingContext context)
        {
            FirstName = (string)info.GetValue("FirstName", typeof(string));
            LastName = (string)info.GetValue("LastName", typeof(string));
            Email = (string)info.GetValue("Email", typeof(string));
            Id = (Guid)info.GetValue("Id", typeof(Guid));
        }

        void AccessLevelChanged(object sender, EventArgs e)
        {
            if (e is AccessLevelChangedEventArgs)
            {
                Debug.WriteLine("Access level changed from: " + (e as AccessLevelChangedEventArgs)?.OldValue + " to: " + (e as AccessLevelChangedEventArgs)?.NewValue);
            }

            if (sender is ExcelUser)
            {
                //do stuff
            }
           
        }

        public override string ToString()
        {
            string s = FirstName ?? "[Empty]";
            s += " ";
            s += LastName ?? "";
            return s;
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info?.AddValue("FirstName", FirstName);
            info?.AddValue("LastName", LastName);
            info?.AddValue("Id", Id);
            info?.AddValue("Email", Email);
        }
    }    

    public class AccessLevelChangedEventArgs:EventArgs
    {
        public UserAccessLevel OldValue { get; set; }
        public UserAccessLevel NewValue { get; set; }
    }
    public class GenericEventArgs : EventArgs
    {
        public object OldValue { get; set; }
        public object NewValue { get; set; }
    }
}
