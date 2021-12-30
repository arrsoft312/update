using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

class MemoryMap {
    private List<MemoryChunk> memoryChunkVector = new List<MemoryChunk>();
    private int loLimitM = 0x7FFFFFFF;
    private int hiLimitM = 0x00000000;
    
    //public int LowestAddress {
        //get { return loLimitM; }
    //}
    //public int HighestAddress {
        //get { return hiLimitM; }
    //}
    public int Size {
        get { return (hiLimitM-loLimitM+1); }
    }
    
    public byte this[int index] {
        get
        {
            int addr = (loLimitM+index);
            foreach (MemoryChunk chunk in memoryChunkVector) {
                if (chunk.In(addr)) {
                    return chunk[addr];
                }
            }
            return 0xFF;
        }
    }
    
    public MemoryMap(string filename) {
        StreamReader inFile;
        try {
            inFile = new StreamReader(filename, Encoding.ASCII, false);
        } catch {
            throw new Exception("Couldn't find '" + filename + "'.");
        }
        
        try {
            int highAddress = 0;
            while (!inFile.EndOfStream) {
                string s = inFile.ReadLine();
                if (s == "") {
                    continue;
                }
                if (s[0] != ':') {
                    throw new Exception("'" + filename + "' is not an Intel HEX formatted file.");
                }
                
                DataBuffer line = s.Remove(0, 1);
                
                byte chksum = 0;
                for (int i = (line.Size-1); i >= 0; i--) {
                    chksum += line[i];
                }
                
                if (chksum != 0) {
                    throw new Exception("'" + filename + "' is damaged.");
                }
                
                byte length = line[0];
                int address = ((line[1] << 8) | line[2]);
                byte type = line[3];
                
                switch (type) {
                    case 0:
                    MemoryChunk chunk = new MemoryChunk((highAddress+address), line.ToByteArray(4, length));
                    for (int i = chunk.StartAddress; i <= chunk.EndAddress; i++) {
                        foreach (MemoryChunk chunk2 in memoryChunkVector) {
                            if (chunk2.In(i)) {
                                throw new Exception("Overlapping memory assignments in file '" + filename + "'.");
                            }
                        }
                    }
                    memoryChunkVector.Add(chunk);
                    if (loLimitM > chunk.StartAddress) {
                        loLimitM = chunk.StartAddress;
                    }
                    if (hiLimitM < chunk.EndAddress) {
                        hiLimitM = chunk.EndAddress;
                    }
                    break;
                    
                    case 1: return;
                    case 2: highAddress = ((line[4] << 12) | (line[5] << 4)); break;
                    case 3: break;
                    case 4: highAddress = ((line[4] << 24) | (line[5] << 16)); break;
                    case 5: break;
                    default: throw new Exception("'" + filename + "' is damaged.");
                }
            }
        } finally {
            inFile.Close();
        }
    }
}
