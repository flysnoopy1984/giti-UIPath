﻿using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Text;

namespace RPA.Core
{
    public class RPAEmail
    {

        public static void Sent_126(List<string> toAddrList,
            string title,
            string body,
            List<string> ccAddrList = null,
            List<string> bccAddrList = null,
            List<string> attachFileList = null,
            string fromAddr = "yy1313admin@126.com")
        {
            MailMessage mailMessage = new MailMessage();
       

            mailMessage.From = new MailAddress(fromAddr);

            foreach (string toAddr in toAddrList)
            {
                mailMessage.To.Add(new MailAddress(toAddr));  //收件人邮箱地址
            }
            if (ccAddrList != null)
            {
                foreach (string toAddr in ccAddrList)
                {
                    mailMessage.CC.Add(new MailAddress(toAddr));
                }
            }
            if (bccAddrList != null)
            {
                foreach (string toAddr in bccAddrList)
                {
                    mailMessage.Bcc.Add(new MailAddress(toAddr));
                }
            }
            if (attachFileList != null)
            {
                foreach (string filePath in attachFileList)
                {
                    mailMessage.Attachments.Add(new Attachment(filePath));
                }
            }
            mailMessage.Subject = title;
            mailMessage.Body = body;
            mailMessage.IsBodyHtml = true;


            SmtpClient client = new SmtpClient();
            client.Host = "smtp.126.com";
            client.Credentials = new System.Net.NetworkCredential("yy1313admin@126.com", "edifier");
            client.Send(mailMessage);
        }


        public static void Sent(
            List<string> toAddrList,
            string title,
            string body,
            List<string> ccAddrList = null,
            List<string> bccAddrList = null,
            List<string> attachFileList = null,
            string fromAddr = "RPA.workflow@giti.com")
        {

            MailMessage mailMessage = new MailMessage();
          

            mailMessage.From = new MailAddress(fromAddr);

            foreach(string toAddr in toAddrList)
            {
                mailMessage.To.Add(new MailAddress(toAddr));  //收件人邮箱地址
            }
            if (ccAddrList != null)
            {
                foreach (string toAddr in ccAddrList)
                {
                    mailMessage.CC.Add(new MailAddress(toAddr));  
                }
            }
            if (bccAddrList != null)
            {
                foreach (string toAddr in bccAddrList)
                {
                    mailMessage.Bcc.Add(new MailAddress(toAddr));
                }
            }
            if(attachFileList != null)
            {
                foreach (string filePath in attachFileList)
                {
                    mailMessage.Attachments.Add(new Attachment(filePath));
                }
            }
            mailMessage.Subject = title;
            mailMessage.Body = body;
            mailMessage.IsBodyHtml = true;


            SmtpClient client = new SmtpClient();
            client.Host = "mail.giti.com";

            client.Send(mailMessage);

        }
    }
}