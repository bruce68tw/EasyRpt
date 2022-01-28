using Base.Enums;
using Base.Models;
using Base.Services;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading.Tasks;

namespace EasyRpt
{
    public class MyService
    {
        public async Task RunAsync()
        {
            const string preLog = "EasyRpt: ";
            await _Log.InfoAsync(preLog + "Start.");

            #region 1.read XpEasyRpt rows
            var info = "";
            var db = new Db();
            var rpts = await db.GetJsonsAsync("select * from dbo.XpEasyRpt where Status=1");
            if (rpts == null)
            {
                info = "No XpEasyRpt Rows";
                goto lab_exit;
            }
            #endregion

            //send reports loop
            var smtp = _Fun.Smtp;
            foreach (var rpt in rpts)
            {
                #region 2.set mailMessage
                var rptName = rpt["Name"].ToString();
                var email = new EmailDto()
                {
                    Subject = rptName,
                    ToUsers = _Str.ToList(rpt["ToEmails"].ToString()),
                    Body = "Hello, please check attached report.",
                };
                var msg = _Email.DtoToMsg(email, smtp);
                #endregion

                //3.sql to Memory Stream docx
                var ms = new MemoryStream();
                var docx = _Excel.FileToMsDocx(_Fun.DirRoot + "EasyRptData/" + rpt["TplFile"].ToString(), ms); //ms <-> docx
                await _Excel.DocxBySqlAsync(rpt["Sql"].ToString(), docx, 1, db);
                docx.Dispose(); //must dispose, or get empty excel !!

                //4.set attachment
                ms.Position = 0;
                var attach = new Attachment(ms, new ContentType(ContentTypeEstr.Excel))
                {
                    Name = rptName + ".xlsx"
                };
                msg.Attachments.Add(attach);

                //5.send email
                await _Email.SendByMsgAsync(msg, smtp);    //sync send for stream attachment !!
                ms.Close(); //close after send email, or get error: cannot access a closed stream !!

                //log result
                await _Log.InfoAsync(preLog + "Send " + rptName);
            }

            #region 6.close db & log
        lab_exit:
            if (db != null)
                await db.DisposeAsync();
            if (info != "")
                await _Log.InfoAsync(preLog + info);

            await _Log.InfoAsync(preLog + "End.");
            #endregion
        }

    }//class
}
