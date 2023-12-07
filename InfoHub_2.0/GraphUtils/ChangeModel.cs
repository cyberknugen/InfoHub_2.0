using Microsoft.Graph.Models;

namespace InfoHub.GraphUtils
{
    public record ChangeModel()
    {
        public string ChangeText;
        public DateTimeOffset TimeOfChange;
        public Identity ResponsibleUser;

        public override string ToString()
        {
            return String.Format("[{0}] :: {1}", TimeOfChange.ToString("MM-dd HH:mm"), ChangeText);
        }
    }
}
