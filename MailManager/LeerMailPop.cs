using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenPop.Pop3;

namespace GetMail
{
    class LeerMailPop
    {

        private void LeerConOpenPop2()
        {

            Pop3Client objPop3Client = new Pop3Client();
            //MessageModel message = new MessageModel();

            //try
            //{

            //    objPop3Client.Connect("user", "password", "mail.server.com");
            //    email.OpenInbox();

            //    while (email.NextEmail())

            //    {
            //        if (email.IsMultipart)
            //        {
            //            IEnumerator enumerator = email.MultipartEnumerator;
            //            while (enumerator.MoveNext())
            //            {
            //                Pop3Component multipart = (Pop3Component)
            //                enumerator.Current;
            //                if (multipart.IsBody)
            //                {
            //                    Console.WriteLine("Multipart body:" +
            //                    multipart.Body);
            //                }
            //                else
            //                {
            //                    Console.WriteLine("Attachment name=" +
            //                    multipart.Name); // ... etc
            //                }
            //            }
            //        }
            //    }

            //    email.CloseConnection();

            //}
            //catch (Pop3LoginException)
            //{
            //    Console.WriteLine("You seem to have a problem logging in!");
            //}

        }


    }
}
