using System;

struct DataBuffer {
    private byte[] hexStringM;
    
    public int Size {
        get { return hexStringM.Length; }
    }
    
    public byte this[int index] {
        get { return hexStringM[index]; }
    }
    
    public byte[] ToByteArray(int startIndex = 0, int length = -1) {
        if (length < 0) {
            length = hexStringM.Length;
        }
        
        byte[] array = new byte[length];
        for (int i = 0; i < length; i++) {
            array[i] = hexStringM[startIndex+i];
        }
        
        return array;
    }
    
    static public implicit operator DataBuffer(string text) {
        return new DataBuffer(text);
    }
    
    public DataBuffer(string text) {
        int length = text.Length;
        
        hexStringM = new byte[length/2];
        
        for (int i = 0; i < length; i++) {
            int c = text[i];
            if (c >= 'A' && c <= 'F') {
                c -= ('A'-10);
            } else if (c >= 'a' && c <= 'f') {
                c -= ('a'-10);
            } else if (c >= '0' && c <= '9') {
                c -= '0';
            } else {
                throw new FormatException("Could not find any recognizable digits.");
            }
            
            if ((i % 2) == 0) {
                hexStringM[i/2] = (byte)(c << 4);
            } else {
                hexStringM[i/2] |= (byte)c;
            }
        }
    }
}
