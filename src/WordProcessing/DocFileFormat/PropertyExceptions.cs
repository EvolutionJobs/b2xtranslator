using System;
using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;

namespace b2xtranslator.DocFileFormat
{
    public class PropertyExceptions : IVisitable
    {
        /// <summary>
        /// A list of the sprms that encode the differences between 
        /// CHP for a character and the PAP for the paragraph style used.
        /// </summary>
        public List<SinglePropertyModifier> grpprl;

        public PropertyExceptions()
        {
            this.grpprl = new List<SinglePropertyModifier>();
        }

        public PropertyExceptions(byte[] bytes)
        {
            this.grpprl = new List<SinglePropertyModifier>();

            if (bytes.Length != 0)
            {
                //read the sprms
                
                int sprmStart = 0;
                bool goOn = true;
                while (goOn)
                {
                    //enough bytes to read?
                    if (sprmStart + 2 < bytes.Length)
                    {
                        //make spra
                        var opCode = (SinglePropertyModifier.OperationCode)System.BitConverter.ToUInt16(bytes, sprmStart);
                        byte spra = (byte)((int)opCode >> 13);

                        // get size of operand
                        short opSize = (short)SinglePropertyModifier.GetOperandSize(spra);
                        byte lenByte = 0;

                        //operand has variable size
                        if (opSize == 255)
                        {
                            //some opCode need special treatment
                            switch (opCode)
                            {
                                case SinglePropertyModifier.OperationCode.sprmTDefTable:
                                case SinglePropertyModifier.OperationCode.sprmTDefTable10:
                                    //The opSize of the table definition is stored in 2 bytes instead of 1
                                    lenByte = 2;
                                    opSize = System.BitConverter.ToInt16(bytes, sprmStart + 2);
                                    //Word adds an additional byte to the opSize to compensate the additional
                                    //byte needed for the length
                                    opSize--;
                                    break;
                                case SinglePropertyModifier.OperationCode.sprmPChgTabs:
                                    //The tab operand can be bigger than 255 bytes (length byte is set to 255).
                                    //In this case a special calculation of the opSize is needed
                                    lenByte = 1;
                                    opSize = bytes[sprmStart + 2];
                                    if (opSize == 255)
                                    {
                                        byte itbdDelMax = bytes[sprmStart + 3];
                                        byte itbdAddMax = bytes[sprmStart + 3 + 2 * itbdDelMax];
                                        opSize = (short)((itbdDelMax * 4 + itbdAddMax * 3) - 1);
                                    }
                                    break;
                                default:
                                    //The variable length stand in the byte after the opcode
                                    lenByte = 1;
                                    opSize = bytes[sprmStart + 2];
                                    break;
                            }
                        }

                        //copy sprm to array
                        //length is 2byte for the opCode, lenByte for the length, opSize for the length of the operand
                        var sprmBytes = new byte[2 + lenByte + opSize];

                        if (bytes.Length >= sprmStart + sprmBytes.Length)
                        {
                            Array.Copy(bytes, sprmStart, sprmBytes, 0, sprmBytes.Length);

                            //parse
                            var sprm = new SinglePropertyModifier(sprmBytes);
                            this.grpprl.Add(sprm);

                            sprmStart += sprmBytes.Length;
                        }
                        else
                        {
                            goOn = false;
                        }
                    }
                    else
                    {
                        goOn = false;
                    }
                }
            }
        }

        #region IVisitable Members

        public virtual void Convert<T>(T mapping)
        {
            ((IMapping<PropertyExceptions>)mapping).Apply(this);
        }

        #endregion
    }

}
