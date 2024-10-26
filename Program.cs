using OfficeOpenXml;

namespace WFormProjEstimateApp1
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();

            // �ӷ��ɮ׸��|�w�]���Ū�
            string? sourceExcelFilePath = null;
            // �ˬd�O�_���ǻ��Ѽ�
            if (args.Length > 0)
            {
                // �u���Ĥ@�ӰѼ� args[0]�A���ˬd�ɮ׸��|�H�T�O�ɮ׽T��s�b
                if (File.Exists(args[0]))
                {
                    // �ɮ׽T��s�b���ܴN�� sourceExcelFilePath �@���ΨӶ}���ɮת��ӷ� Excel ���| sourceExcelFilePath
                    sourceExcelFilePath = args[0];
                }
            }

            // �N�ɮ׸��|�ǤJ WinForm �Ұ����ε{���C�p�G�S���ɮ׸��|�w�]���Ū�
            Application.Run(new WFormProjEstimate(sourceExcelFilePath));     
        }
    }
}