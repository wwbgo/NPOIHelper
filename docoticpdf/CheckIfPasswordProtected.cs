using System;
using System.Text;
using BitMiracle.Docotic.Pdf;

namespace docoticpdf {
    public static class CheckIfPasswordProtected {
        public static void Main() {
            //CheckPasswordProtected();
            try {
                CheckPasswordProtected();
            } catch (Exception ex) {
                Console.WriteLine(ex);
            }
            Console.Read();
        }

        private static void CheckPasswordProtected() {
            var message = new StringBuilder();

            string[] documentsToCheck = { "chizi_a4_zaixianchizi.pdf1","Sample Documentation.doc.pdf" };
            foreach (var fileName in documentsToCheck) {
                var passwordProtected = PdfDocument.IsPasswordProtected(@"E:\Doc\" + fileName);
                if (passwordProtected)
                    message.AppendFormat("{0} - REQUIRES PASSWORD\r\n",fileName);
                else
                    message.AppendFormat("{0} - DOESN'T REQUIRE PASSWORD\r\n",fileName);
            }

            Console.Write(message.ToString());
            Console.Read();
        }
    }
}
