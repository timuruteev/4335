using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_4435
{
    public class serDeser
    {
        public int Id { get; set; }
        public string CodeOrder { get; set; }
        public string CreateDate { get; set; }
        public string CreateTime { get; set; }
        public string CodeClient { get; set; }
        public string Services { get; set; }
        public string Status { get; set; }
        public string ClosedDate { get; set; }
        public string ProkatTime { get; set; }

        public serDeser()
        {

        }

        public serDeser(int id, string codeOrder, string createDate, string createTime, string codeClient, string services, string status, string closedDate, string prokatTime)
        {
            Id = id;
            CodeOrder = codeOrder;
            CreateDate = createDate;
            CreateTime = createTime;
            CodeClient = codeClient;
            Services = services;
            Status = status;
            ClosedDate = closedDate;
            ProkatTime = prokatTime;
        }
    }
}
