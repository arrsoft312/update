using System;

struct MemoryChunk {
    private byte[] hexStringM;
    private int addressM;
    
    public int StartAddress {
        get { return addressM; }
    }
    public int EndAddress {
        get { return (addressM + hexStringM.Length - 1); }
    }
    
    public byte this[int addr] {
        get { return hexStringM[addr-addressM]; }
    }
    
    public bool In(int addr) {
        return (addr >= addressM && addr < (addressM + hexStringM.Length));
    }
    
    public MemoryChunk(int addr, byte[] array) {
        int length = array.Length;
        
        hexStringM = new byte[length];
        addressM = addr;
        
        for (int i = 0; i < length; i++) {
            hexStringM[i] = array[i];
        }
    }
}
