using System;
using System.Net;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Diagnostics ;

namespace WinSMTPTest
{
	/// <summary>
	/// Summary description for ClassTCP.
	/// </summary>
	public class ClassTCP
	{
		private string COPY_INFO;
		private TcpClient tcpClient;
		private NetworkStream stream;
		private string server;
		private int port;
		private string serverManifest;	//risposta del server dopo la connessione

		private bool readResponseMessage;	//indica quando si è nel comando data, non deve leggere la risposta
		private bool isConnect;

		public ClassTCP(string footer)
		{
			server = "127.0.0.1";
			port = 25;
			readResponseMessage = true;
			serverManifest = "";
			isConnect = false;
			COPY_INFO = footer;
		}
		public ClassTCP(string server, int port)
		{
			this.server = server;
			this.port = port;
		}

		public string Open(string server, string port) 
		{
			this.server = server;
			try 
			{
				this.port = Convert.ToInt32(port);
			}
			catch
			{
				this.port = 25;
			}
		return Open();
	}
		public string Open()
		{
			try 
			{
				tcpClient = new TcpClient(server, port);
				//tcpClient.NoDelay = true;
				tcpClient.SendTimeout = 60000;
				tcpClient.ReceiveTimeout = 60000;
				stream = tcpClient.GetStream();
				readResponseMessage = true;
				serverManifest = GetResponse();
				isConnect = true;
				return serverManifest;
			}
			catch(SocketException ex)
			{
				isConnect = false;
				return "";
			}

		}

		public void Close()
		{
			isConnect = false;
			try 
			{
				tcpClient.Close();
			}
			catch(Exception ex)
			{
			}
		}

		public string ServerManifest
		{
			get 
			{
				return serverManifest;
			}
		}

		public bool IsConnect
		{
			get 
			{
				return isConnect;
			}
		}

		public string GetResponse()
		{
			string line = "";
			string r = "";

			try 
			{
				StreamReader sr = new StreamReader(stream, Encoding.ASCII);
				do 
				{
					line = sr.ReadLine();
					r = String.Concat(r, line, "\r\n");
				}while (line.Substring(3, 1).CompareTo("-") == 0);
			}
			catch(Exception e)
			{
				r += String.Concat("\r\n", e.ToString());
			}
			return r;
		}

		public string SendCommand(string cmd)
		{
			string r = "";
			string cmdLower = cmd.ToLower();

			if ( readResponseMessage == false && cmdLower.CompareTo(".") == 0) 
			{
				//inserisco il copyright
				SendCommand(COPY_INFO);
				readResponseMessage = true;
			}

			if (cmd.Length > 0 || readResponseMessage == false)
			{
				// Send the message to the connected TcpServer. 
				try 
				{
					Byte[] data = System.Text.Encoding.ASCII.GetBytes(cmd + "\r\n");         
					stream.Write(data, 0, data.Length);
				
					if (readResponseMessage == true)
						r = GetResponse();

					if (cmdLower.CompareTo("data") == 0 
						|| cmdLower.StartsWith("data ") == true ||	cmdLower.StartsWith("data\t" ) == true) 
						readResponseMessage=false;
	
					if (cmdLower.CompareTo("quit") == 0 
						|| cmdLower.StartsWith("quit ") == true || cmdLower.StartsWith("quit\t") == true)
						Close();
				}
				catch (Exception e)
				{
					r += e.Message;
				}
			}
			return r;
		}
	}
}
