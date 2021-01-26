using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using System.Text;

namespace RPA.Core
{
    public class BaseTask
    {
        public static void SendMail()
        {
            string host = "smtp.exmail.qq.com";// 
            string userName = "test123@gititire.com";// 
            string password = "Welcome1>";// 
            int port = 465;

            SmtpClient client = new SmtpClient();
            client.DeliveryMethod = SmtpDeliveryMethod.Network;//指定电子邮件发送方式    
            client.Host = host;//邮件服务器
                               //    client.Port = port;
            client.UseDefaultCredentials = true;
            client.Credentials = new System.Net.NetworkCredential(userName, password);//用户名、密码

            //////////////////////////////////////
            string strfrom = userName;
            string strto = "song.fuwei@giti.com";
            string strcc = "song.fuwei@giti.com";//抄送


            string subject = "这是测试邮件标题5";//邮件的主题             
            string body = "测试邮件内容5";//发送的邮件正文  

            System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
            msg.From = new MailAddress(strfrom, "xyf");
            msg.To.Add(strto);
            msg.CC.Add(strcc);

            msg.Subject = subject;//邮件标题   
            msg.Body = body;//邮件内容   
            msg.BodyEncoding = System.Text.Encoding.UTF8;//邮件内容编码   
            msg.IsBodyHtml = true;//是否是HTML邮件   
            msg.Priority = MailPriority.High;//邮件优先级   


            try
            {
                client.Send(msg);
                Console.WriteLine("发送成功");
            }
            catch (System.Net.Mail.SmtpException ex)
            {
                Console.WriteLine(ex.Message, "发送邮件出错");
            }
        }

        public static void KillProcess(string processName)
        {
            foreach (Process p in Process.GetProcesses())
            {
             //   Console.WriteLine(p.ProcessName);
                if (p.ProcessName.Contains(processName))
                {
                    try
                    {
                        p.Kill();
                        p.WaitForExit(); // possibly with a timeout
                        Console.WriteLine($"已杀掉{processName}进程！！！");
                    }
                    catch (Win32Exception e)
                    {
                        Console.WriteLine(e.Message.ToString());
                    }
                    catch (InvalidOperationException e)
                    {
                        Console.WriteLine(e.Message.ToString());
                    }
                }

            }
        }
    }
}
