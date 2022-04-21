using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BotManager
{
    internal class botUser
    {
        long id;

        internal byte step { get {return step; } set {; } }
        
        
        internal botUser (long id)
        {
            this.id = id;
            step = 1;
        }



    }
}
