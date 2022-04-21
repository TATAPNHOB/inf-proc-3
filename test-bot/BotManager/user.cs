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

        internal byte step;

        internal byte tries;

        internal DateTime banDate;

        internal botUser (long id)
        {
            this.id = id;
            step = 1;
        }



    }
}
