using System;
using System.Windows.Forms;

namespace InvAddIn
{
    internal class ExceptionHandler
    {
        public static void HandleException(Exception ex)
        {
            MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Console.WriteLine($"Error: {ex.Message}");
        }
        //Add log.txt file (logger)
    }
}
