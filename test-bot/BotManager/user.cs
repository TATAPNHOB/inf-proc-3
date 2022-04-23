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
        public string email;
        //5 минут
        TimeSpan dif = new TimeSpan(0,5,0);

        internal byte step;

        internal byte tries;

        internal DateTime banDate;

        //проверка, в бане ли челик
        internal bool IsBanned { get {return DateTime.Now.Subtract(banDate).CompareTo(new TimeSpan(0,5*tries/5,0))<=0; } }

        internal botUser (long id)
        {
            this.id = id;
            step = 1;
            tries = 0;
        }



    }
}
