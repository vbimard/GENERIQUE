using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//specific toast
using Windows.UI.Notifications;
using Windows.Data.Xml.Dom;

namespace AF_Toast_Notification
{
    public class Toast_Notification
    {
        private string ToastTitle;
        private string ToastMessage;
        private string ToastImagePath;
        private int ToastType;

        public Toast_Notification(string title, string message, string imagePath, int type)
        {
            ToastTitle = title;
            ToastMessage = message;
            ToastImagePath = imagePath;
            ToastType = type;
        }
        public void CreateNotification0()
        {
            // template ref : https://docs.microsoft.com/en-us/previous-versions/windows/apps/hh761494(v=win.10)
            // autre ref https://books.google.fr/books?id=D7K5R74XKNQC&pg=PA143&lpg=PA143&dq=toast+textandimage+sample+c%23&source=bl&ots=FVzVv8jvML&sig=ACfU3U1VlrtjgabL-_stjxA_RfOf9LzflA&hl=fr&sa=X&ved=2ahUKEwj5h-zd2p3gAhXiBWMBHfkEBrMQ6AEwEnoECAAQAQ#v=onepage&q=toast%20textandimage%20sample%20c%23&f=false
            string Toast = "";
            switch (ToastType)
            {
                case 20:
               

                    Toast = "<toast><visual><binding template=\"ToastImageAndText02\">";
                    Toast += "< image id = \"1\" >src=\"image1\" alt=\"image1\">";
                    Toast += "< text id = \"1\" >" + ToastTitle + "</text>";
                    Toast += "< text id = \"2\" >" + ToastMessage + "</text>";
                    Toast += "</binding></visual></toast>";
                    break;

                case 2:
                    {
                        //Basic Toast
                        /*
                           <toast>
                            <visual>
                                <binding template="ToastText02">
                                    <text id="1">headlineText</text>
                                    <text id="2">bodyText</text>
                                </binding>  
                            </visual>
                            </toast>
                         
                         */
                        Toast = "<toast><visual><binding template=\"ToastImageAndText02\">";
                        Toast += "< text id = \"1\" >" + ToastTitle + "</text>";
                        Toast += "< text id = \"2\" >" + ToastMessage + "</text>";
                        Toast += "</binding></visual></toast>";
                        break;
                    }

                case 1:
                    {
                       

                        //Basic Toast
                        Toast = "<toast><visual><binding template=\"ToastImageAndText01\">";
                        Toast += "< text id = \"1\" >" + ToastMessage + "</text>";
                        Toast += "</binding></visual></toast>";
                        break;
                    }
                default:
                    {
                        Toast = "<toast><visual><binding template=\"ToastImageAndText01\"><text id = \"1\" >";
                        Toast += "EveryThing is ok on this toast";
                        Toast += "</text></binding></visual></toast>";
                        break;
                    }
            }

            XmlDocument tileXml = new XmlDocument();
            tileXml.LoadXml(Toast);
            var toast = new ToastNotification(tileXml);
            ToastNotificationManager.CreateToastNotifier("AlmaCam").Show(toast);
        }
        public bool CreateNotification()
        {
            try
            {
                bool rst = false;
                if (ToastType == 2) { 
                XmlDocument template = ToastNotificationManager.GetTemplateContent(ToastTemplateType.ToastText02);
                IXmlNode mainnode = template.SelectSingleNode("/toast");
                XmlNodeList text = template.GetElementsByTagName("text");
                text[0].AppendChild(template.CreateTextNode("title"));
                text[1].AppendChild(template.CreateTextNode("lorem ipsum"));
                ToastNotification toast = new ToastNotification(template);
                ToastNotificationManager.CreateToastNotifier().Show(toast);
                    rst = true;
                }
                return rst;

            }
            catch (Exception ie)
            {
                System.Windows.Forms.MessageBox.Show(System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ie.Message);
                return false;
            }

            finally
            { }
        }


    }
}
