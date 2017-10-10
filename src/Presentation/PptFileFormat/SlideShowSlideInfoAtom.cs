/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecord(1017)]
    public class SlideShowSlideInfoAtom: Record
    {
        public int slideTime;
        public uint soundIdRef;
        public byte effectDirection;
        public byte effectType;
        public bool fManualAdvance;
        public bool fHidden;
        public bool fSound;
        public bool fLoopSound;
        public bool fStopSound;
        public bool fAutoAdvance;
        public bool fCursorVisible;
        public byte speed;


        public SlideShowSlideInfoAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            slideTime = this.Reader.ReadInt32();
            soundIdRef = this.Reader.ReadUInt32();
            effectDirection = this.Reader.ReadByte();
            effectType = this.Reader.ReadByte();

            int flags = this.Reader.ReadInt16();

            fManualAdvance = Tools.Utils.BitmaskToBool(flags, 0x1);
            fHidden = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);
            fSound = Tools.Utils.BitmaskToBool(flags, 0x1 << 4);
            fLoopSound = Tools.Utils.BitmaskToBool(flags, 0x1 << 6);
            fStopSound = Tools.Utils.BitmaskToBool(flags, 0x1 << 8);
            fAutoAdvance = Tools.Utils.BitmaskToBool(flags, 0x1 << 10);
            fCursorVisible = Tools.Utils.BitmaskToBool(flags, 0x1 << 12);

            speed = this.Reader.ReadByte();

            this.Reader.BaseStream.Position = this.Reader.BaseStream.Length;
        }

    }

    [OfficeRecord(12001)]
    public class Comment10Atom : Record
    {
        public Comment10Atom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(12011)]
    public class SlideTime10Atom : Record
    {
        public byte[] fileTime;
        public SlideTime10Atom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            fileTime = this.Reader.ReadBytes(8);
        }
    }

    [OfficeRecord(61764)]
    public class ExtTimeNodeContainer : RegularContainer
    {
        public ExtTimeNodeContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(61735)]
    public class TimeNodeAtom : Record
    {
        public uint restart;
        public TimeNodeTypeEnum type;
        public uint fill;
        public int duration;
        public bool fFillProperty;
        public bool fRestartProperty;
        public bool fGroupingTypeProperty;
        public bool fDurationProperty;
        public TimeNodeAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Reader.ReadBytes(4); //reserved
            restart = this.Reader.ReadUInt32();
            type = (TimeNodeTypeEnum)this.Reader.ReadInt32();
            fill = this.Reader.ReadUInt32();

            //according to the spec there would be 5 unused bytes
            //but in reality it is 8
            this.Reader.ReadBytes(8); //reserved

            duration = this.Reader.ReadInt32();
                      
            int flags = this.Reader.ReadInt32();
            fFillProperty = Tools.Utils.BitmaskToBool(flags, 0x1 << 0);
            fRestartProperty = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);

            bool dummy = Tools.Utils.BitmaskToBool(flags, 0x1 << 2); 

            fGroupingTypeProperty = Tools.Utils.BitmaskToBool(flags, 0x1 << 3);
            fDurationProperty = Tools.Utils.BitmaskToBool(flags, 0x1 << 4);

            this.Reader.BaseStream.Position = this.Reader.BaseStream.Length;
        }

        public enum TimeNodeTypeEnum
        {
            parallel = 0x0,
            sequential,
            behavior,
            media
        }
    }

    [OfficeRecord(61757)]
    public class TimePropertyList4TimeNodeContainer : RegularContainer
    {
        public TimePropertyList4TimeNodeContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(11003)]
    public class VisualShapeAtom : Record //can be VisualSoundAtom, VisualShapeChartElementAtom, VisualShapeGeneralAtom
    {
        public TimeVisualElementEnum type;
        public ElementTypeEnum refType;
        public uint shapeIdRef;
        public int data1;
        public int data2;

        public VisualShapeAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            type = (TimeVisualElementEnum)this.Reader.ReadInt32();
            refType = (ElementTypeEnum)this.Reader.ReadInt32();
            shapeIdRef = this.Reader.ReadUInt32();
            data1 = this.Reader.ReadInt32();
            data2 = this.Reader.ReadInt32();
        }
    }

    public enum ElementTypeEnum
    {
        ShapeType = 1,
        SoundType = 2
    }

    [OfficeRecord(11008)]
    public class HashCode10Atom : Record
    {
        public uint hash;
        public HashCode10Atom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            hash = this.Reader.ReadUInt32();
        }
    }

    [OfficeRecord(11009)]
    public class VisualPageAtom : Record
    {
        public TimeVisualElementEnum type;
        public VisualPageAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            type = (TimeVisualElementEnum)this.Reader.ReadInt32();
        }
    }

    public enum TimeVisualElementEnum
    {
        Shape = 0,
        Page,
        TextRange,
        Audio,
        Video,
        ChartElement,
        ShapeOnly,
        AllTextRange
    }

    [OfficeRecord(11010)]
    public class BuildListContainer : RegularContainer
    {
        public BuildListContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(61733)]
    public class TimeConditionContainer : RegularContainer
    {
        public TimeConditionContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(61762)]
    public class TimeVariantValue : Record //can be TimeVariantBool, TimeVariantInt, TimeVariantFloat or TimeVariantString
    {
        public TimeVariantTypeEnum type;
        public int? intValue = null;
        public bool? boolValue = null;
        public float? floatValue = null;
        public string stringValue = null;
        public TimeVariantValue(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            type = (TimeVariantTypeEnum)this.Reader.ReadByte();
            switch (type)
            {
                case TimeVariantTypeEnum.Bool:
                    boolValue = this.Reader.ReadBoolean();
                    break;
                case TimeVariantTypeEnum.Float:
                    floatValue = this.Reader.ReadSingle();
                    break;
                case TimeVariantTypeEnum.Int:
                    intValue = this.Reader.ReadInt32();
                    break;
                case TimeVariantTypeEnum.String:
                    stringValue = Encoding.Unicode.GetString(this.Reader.ReadBytes((int)size - 1));
                    stringValue = stringValue.Replace("\0", "");
                    break;
            }
        }
    }

    public enum TimePropertyID4TimeNode
    {
        Display = 0x2,
        MasterPos = 0x5,
        SlaveType = 0x6,
        EffectID = 0x9,
        EffectDir = 0xA,
        EffectType = 0xB,
        AfterEffect = 0xD,
        SlideCount = 0xF,
        TimeFilter = 0x10,
        EventFilter = 0x11,
        HideWhenStopped = 0x12,
        GroupID = 0x13,
        EffectNodeType = 0x14,
        PlaceholderNode = 0x15,
        MediaVolume = 0x16,
        MediaMute = 0x17,
        ZoomToFullScreen = 0x1A
    }

    public enum TimeVariantTypeEnum
    {
        Bool = 0,
        Int,
        Float,
        String
    }

    [OfficeRecord(61737)]
    public class TimeModifierAtom : Record 
    {
        public uint type;
        //0 repeat count
        //1 repeat duration
        //2 speed
        //3 accelerate
        //4 decelerate
        //5 automatic reverse

        public uint value;
        public TimeModifierAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            type = this.Reader.ReadUInt32();
            value = this.Reader.ReadUInt32();
        }
    }

    [OfficeRecord(61739)]
    public class TimeAnimateBehaviourContainer : RegularContainer
    {
        public TimeAnimateBehaviourContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(61745)]
    public class TimeSetBehaviourContainer : RegularContainer
    {
        public TimeSetBehaviourContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(61736)]
    public class TimeConditionAtom : Record
    {
        public TriggerObjectEnum triggerObject;
        public uint triggerEvent;
        public uint id;
        public int delay;
        public TimeConditionAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            triggerObject = (TriggerObjectEnum)this.Reader.ReadInt32();
            triggerEvent = this.Reader.ReadUInt32();
            id = this.Reader.ReadUInt32();
            delay = this.Reader.ReadInt32();
        }

        public enum TriggerObjectEnum
        {
            None = 0,
            VisualElement,
            TimeNode,
            RuntimeNodeRef
        }
    }

    [OfficeRecord(61748)]
    public class TimeAnimateBehaviorAtom : Record
    {
        public uint calcMode;
        //0 discrete
        //1 linear
        //2 formula

        public bool fByPropertyUsed;
        public bool fFromPropertyUsed;
        public bool fToPropertyUsed;
        public bool fCalcModePropertyUsed;
        public bool fAnimationValuesPropertyUsed;
        public bool fValueTypePropertyUsed;
        public TimeAnimateBehaviorValueTypeEnum valueType;

        public TimeAnimateBehaviorAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            calcMode = this.Reader.ReadUInt32();
            int flags = this.Reader.ReadInt32();

            fByPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1);
            fFromPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);
            fToPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);
            fCalcModePropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 3);
            fAnimationValuesPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 4);
            fValueTypePropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 5);
            valueType = (TimeAnimateBehaviorValueTypeEnum)this.Reader.ReadInt32();
        }        
    }

    public enum TimeAnimateBehaviorValueTypeEnum
    {
        String = 0,
        Number,
        Color
    }

    [OfficeRecord(61759)]
    public class TimeAnimationValueListContainer : RegularContainer
    {
        public TimeAnimationValueListContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(61738)]
    public class TimeBehaviorContainer : RegularContainer
    {
        public TimeBehaviorContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(61747)]
    public class TimeBehaviorAtom : Record
    {
        public bool fAdditivePropertyUsed;
        public bool fAttributeNamesPropertyUsed;
        public uint behaviorAdditive;
        public uint behaviorAccumulate;
        public uint behaviorTransform;

        public TimeBehaviorAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            int flags = this.Reader.ReadInt32();

            fAdditivePropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1);
            fAttributeNamesPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);

            behaviorAdditive = this.Reader.ReadUInt32();
            behaviorAccumulate = this.Reader.ReadUInt32();
            behaviorTransform = this.Reader.ReadUInt32();
        }
    }

    [OfficeRecord(61758)]
    public class TimeStringListContainer : RegularContainer
    {
        public TimeStringListContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(61756)]
    public class ClientVisualElementContainer : RegularContainer
    {
        public ClientVisualElementContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(61763)]
    public class TimeAnimationValueAtom : Record
    {
        public int time;
        public TimeAnimationValueAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            time = this.Reader.ReadInt32();
        }
    }

    [OfficeRecord(61754)]
    public class TimeSetBehaviorAtom : Record
    {
        public bool fToPropertyUsed;
        public bool fValueTypePropertyUsed;
        public TimeAnimateBehaviorValueTypeEnum valueType;
 
        public TimeSetBehaviorAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            int flags = this.Reader.ReadInt32();

            fToPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1);
            fValueTypePropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);

            valueType = (TimeAnimateBehaviorValueTypeEnum)this.Reader.ReadInt32();
        }
    }

    [OfficeRecord(61761)]
    public class TimeSequenceDataAtom : Record
    {
        public uint concurrency;
        public uint nextAction;
        public uint previousAction;
        public bool fConcurrencyPropertyUsed;
        public bool fNextActionPropertyUsed;
        public bool fPreviousActionPropertyUsed;

        public TimeSequenceDataAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            concurrency = this.Reader.ReadUInt32();
            nextAction = this.Reader.ReadUInt32();
            previousAction = this.Reader.ReadUInt32();
            this.Reader.ReadInt32(); //reserved

            int flags = this.Reader.ReadInt32();
            fConcurrencyPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1);
            fNextActionPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);
            fPreviousActionPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);

        }
    }

    [OfficeRecord(61742)]
    public class TimeMotionBehaviorContainer : RegularContainer
    {
        public TimeMotionBehaviorContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(61751)]
    public class TimeMotionBehaviorAtom : Record
    {
        public bool fByPropertyUsed;
        public bool fFromPropertyUsed;
        public bool fToPropertyUsed;
        public bool fOriginPropertyUsed;
        public bool fPathPropertyUsed;
        public bool fEditRotationPropertyUsed;
        public bool fPointsTypesPropertyUsed;
        public float fXBy;
        public float fYBy;
        public float fXFrom;
        public float fYFrom;
        public float fXTo;
        public float fYTo;
        public uint behaviorOrigin;

        public TimeMotionBehaviorAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            int flags = this.Reader.ReadInt32();
            fByPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1);
            fFromPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);
            fToPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);
            fOriginPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 3);
            fPathPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 4);
            // 1 bit reserved
            fEditRotationPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 6);           
            fPointsTypesPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 7);
            fXBy = this.Reader.ReadSingle();
            fYBy = this.Reader.ReadSingle();
            fXFrom = this.Reader.ReadSingle();
            fYFrom = this.Reader.ReadSingle();
            fXTo = this.Reader.ReadSingle();
            fYTo = this.Reader.ReadSingle();
            behaviorOrigin = this.Reader.ReadUInt32();
        }
    }


    [OfficeRecord(61744)]
    public class TimeScaleBehaviorContainer : RegularContainer
    {
        public TimeScaleBehaviorContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(61741)]
    public class TimeEffectBehaviorContainer : RegularContainer
    {
        public TimeEffectBehaviorContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(61750)]
    public class TimeEffectBehaviorAtom : Record
    {
        public bool fTransitionPropertyUsed;
        public bool fTypePropertyUsed;
        public bool fProgressPropertyUsed;
        public bool fRuntimeContextObsolete;
        public uint effectTransition;

        public TimeEffectBehaviorAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            int flags = this.Reader.ReadInt32();
            fTransitionPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1);
            fTypePropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);
            fProgressPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);
            fRuntimeContextObsolete = Tools.Utils.BitmaskToBool(flags, 0x1 << 3);
            effectTransition = this.Reader.ReadUInt32();
        }
    }

    [OfficeRecord(61752)]
    public class TimeRotationBehaviorAtom : Record
    {
        public bool fByPropertyUsed;
        public bool fFromPropertyUsed;
        public bool fToPropertyUsed;
        public bool fDirectionPropertyUsed;

        public float fBy;
        public float fFrom;
        public float fTo;

        public uint rotationDirection;

        public TimeRotationBehaviorAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            int flags = this.Reader.ReadInt32();
            fByPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1);
            fFromPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);
            fToPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);
            fDirectionPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 3);

            fBy = this.Reader.ReadSingle();
            fFrom = this.Reader.ReadSingle();
            fTo = this.Reader.ReadSingle();            

            rotationDirection = this.Reader.ReadUInt32();
        }
    }

    [OfficeRecord(61753)]
    public class TimeScaleBehaviorAtom : Record
    {
        public bool fByPropertyUsed;
        public bool fFromPropertyUsed;
        public bool fToPropertyUsed;
        public bool fZoomPropertyUsed;

        public float fXBy;
        public float fYBy;
        public float fXFrom;
        public float fYFrom;
        public float fXTo;
        public float fYTo;
        public byte fZoomContents;

       public TimeScaleBehaviorAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            int flags = this.Reader.ReadInt32();
            fByPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1);
            fFromPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);
            fToPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);
            fZoomPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 3);

            fXBy = this.Reader.ReadSingle();
            fYBy = this.Reader.ReadSingle();
            fXFrom = this.Reader.ReadSingle();
            fYFrom = this.Reader.ReadSingle();
            fXTo = this.Reader.ReadSingle();
            fYTo = this.Reader.ReadSingle();

            fZoomContents = this.Reader.ReadByte();
        }
    }

    [OfficeRecord(61743)]
    public class TimeRotationBehaviorContainer : RegularContainer
    {
        public TimeRotationBehaviorContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
        }
    }


    [OfficeRecord(61746)]
    public class TimeCommandBehaviorContainer : RegularContainer
    {
        public TimeCommandBehaviorContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
        }
    }

    [OfficeRecord(61765)]
    public class SlaveContainer : RegularContainer
    {
        public SlaveContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(61740)]
    public class TimeColorBehaviorContainer : RegularContainer
    {
        public TimeColorBehaviorContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(61749)]
    public class TimeColorBehaviorAtom : Record
    {
        public bool fByPropertyUsed;
        public bool fFromPropertyUsed;
        public bool fToPropertyUsed;
        public bool fColorSpacePropertyUsed;
        public bool fDirectionPropertyUsed;

        public ColorStruct colorBy;
        public ColorStruct colorFrom;
        public ColorStruct colorTo;  

        public TimeColorBehaviorAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            int flags = this.Reader.ReadInt32();
            fByPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1);
            fFromPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);
            fToPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);
            fColorSpacePropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 3);
            fDirectionPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 4);

            colorBy = new ColorStruct(this.Reader.ReadBytes(16));
            colorFrom = new ColorStruct(this.Reader.ReadBytes(16));
            colorTo = new ColorStruct(this.Reader.ReadBytes(16));
        }
    }

    [OfficeRecord(61755)]
    public class TimeCommandBehaviorAtom : Record
    {
        public bool fTypePropertyUsed;
        public bool fCommandPropertyUsed;

        public UInt32 commandBehaviorType;

        public TimeCommandBehaviorAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            int flags = this.Reader.ReadInt32();
            fTypePropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1);
            fCommandPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);

            commandBehaviorType = this.Reader.ReadUInt32();
        }
    }

    [OfficeRecord(61760)]
    public class TimeIterateDataAtom : Record
    {
        public float iterateInterval;
        public uint iterateType;
        public uint iterateDirection;
        public uint iterateIntervalType;

        public bool fIterateDirectionPropertyUsed;
        public bool fIterateTypePropertyUsed;
        public bool fIterateIntervalPropertyUsed;
        public bool fIterateIntervalTypePropertyUsed;


        public TimeIterateDataAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            
            //this value "should" be an unsigned integer according to the spec, but it seems to be a single in reality
            iterateInterval = this.Reader.ReadSingle();

            iterateType = this.Reader.ReadUInt32();
            iterateDirection = this.Reader.ReadUInt32();
            iterateIntervalType = this.Reader.ReadUInt32();

            int flags = this.Reader.ReadInt32();
            fIterateDirectionPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1);
            fIterateTypePropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);
            fIterateIntervalPropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);
            fIterateIntervalTypePropertyUsed = Tools.Utils.BitmaskToBool(flags, 0x1 << 3);
        }
    }

    public struct ColorStruct
    {
        public uint model;
        public int val1;
        public int val2;
        public int val3;

        public ColorStruct(byte[] data)
        {
            model = BitConverter.ToUInt32(data, 0);
            val1 = BitConverter.ToInt32(data, 4);
            val2 = BitConverter.ToInt32(data, 8);
            val3 = BitConverter.ToInt32(data, 12);
        }
    }

}
