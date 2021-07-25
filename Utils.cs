using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Kavod.Vba.Compression;

namespace BadAssMacros

{
    class Utils
    {
		public static Dictionary<string, string> ParseArgs(string[] args)
		{
			Dictionary<string, string> ret = new Dictionary<string, string>();
			if (args.Length > 0)
			{
				for (int i = 0; i < args.Length; i += 2)
				{
					if (args[i].Substring(1).ToLower() == "l")
					{
						ret.Add(args[i].Substring(1).ToLower(), "true");
					}
					else
					{
						ret.Add(args[i].Substring(1).ToLower(), args[i + 1]);
					}
				}
			}
			return ret;
		}
		
		public static List<ModuleInformation> ParseModulesFromDirStream(byte[] dirStream)
		{
			// 2.3.4.2 dir Stream: Version Independent Project Information
			// https://msdn.microsoft.com/en-us/library/dd906362(v=office.12).aspx
			// Dir stream is ALWAYS in little endian

			List<ModuleInformation> modules = new List<ModuleInformation>();

			int offset = 0;
			UInt16 tag;
			UInt32 wLength;
			ModuleInformation currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };

			while (offset < dirStream.Length)
			{
				tag = GetWord(dirStream, offset);
				wLength = GetDoubleWord(dirStream, offset + 2);

				// taken from Pcodedmp
				if (tag == 9)
					wLength = 6;
				else if (tag == 3)
					wLength = 2;

				switch (tag)
				{
					// MODULESTREAMNAME Record
					case 26:
						currentModule.moduleName = System.Text.Encoding.UTF8.GetString(dirStream, (int)offset + 6, (int)wLength);
						break;

					// MODULEOFFSET Record
					case 49:
						currentModule.textOffset = GetDoubleWord(dirStream, offset + 6);
						modules.Add(currentModule);
						currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };
						break;
				}

				offset += 6;
				offset += (int)wLength;
			}

			return modules;
		}

		public class ModuleInformation
		{
			// Name of VBA module stream
			public string moduleName;

			// Offset of VBA CompressedSourceCode in VBA module stream
			public UInt32 textOffset;
		}

		public static UInt16 GetWord(byte[] buffer, int offset)
		{
			var rawBytes = new byte[2];
			Array.Copy(buffer, offset, rawBytes, 0, 2);
			return BitConverter.ToUInt16(rawBytes, 0);
		}

		public static UInt32 GetDoubleWord(byte[] buffer, int offset)
		{
			var rawBytes = new byte[4];
			Array.Copy(buffer, offset, rawBytes, 0, 4);
			return BitConverter.ToUInt32(rawBytes, 0);
		}
		public static byte[] Compress(byte[] data)
		{
			var buffer = new DecompressedBuffer(data);
			var container = new CompressedContainer(buffer);
			return container.SerializeData();
		}
		public static byte[] Decompress(byte[] data)
		{
			var container = new CompressedContainer(data);
			var buffer = new DecompressedBuffer(container);
			return buffer.Data;
		}
		public static string GetVBATextFromModuleStream(byte[] moduleStream, UInt32 textOffset)
		{
			string vbaModuleText = Encoding.UTF8.GetString(Decompress(moduleStream.Skip((int)textOffset).ToArray()));
			return vbaModuleText;
		}
		public static byte[] RemovePcodeInModuleStream(byte[] moduleStream, UInt32 textOffset, string OG_VBACode)
		{
			return Compress(Encoding.UTF8.GetBytes(OG_VBACode)).ToArray();
		}
		public static string GetOutFilename(String filename)
		{
			string fn = Path.GetFileNameWithoutExtension(filename);
			string ext = Path.GetExtension(filename);
			string path = Path.GetDirectoryName(filename);
			return Path.Combine(path, fn + ext);
		}
		public static byte[] HexToByte(string hex)
		{
			hex = hex.Replace("-", "");
			byte[] raw = new byte[hex.Length / 2];
			for (int i = 0; i < raw.Length; i++)
			{
				raw[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
			}
			return raw;
		}
		public static byte[] ChangeOffset(byte[] dirStream)
		{
			int offset = 0;
			UInt16 tag;
			UInt32 wLength;

			// Change MODULEOFFSET to 0
			string zeros = "\0\0\0\0";
			
			while (offset < dirStream.Length)
			{
				tag = GetWord(dirStream, offset);
				wLength = GetDoubleWord(dirStream, offset + 2);

				// taken from Pcodedmp
				if (tag == 9)
					wLength = 6;
				else if (tag == 3)
					wLength = 2;

				switch (tag)
				{
					// MODULEOFFSET Record
					case 49:
						uint offset_change = GetDoubleWord(dirStream, offset + 6);
						UTF8Encoding encoding = new UTF8Encoding();
						encoding.GetBytes(zeros, 0, (int)wLength, dirStream, (int)offset + 6);
						break;
				}

				offset += 6;
				offset += (int)wLength;
			}
			return dirStream;
		}
	}
}
