﻿using System;
using System.Runtime.Serialization;

namespace UiPathTeam.Office.Comments
{
    public class CustomException : Exception
    {
        public CustomException()
        {}

        public CustomException(string message) : base(message)
        {}

        public CustomException(string message, Exception inner) : base(message, inner)
        {}

        public CustomException(SerializationInfo serializationInfo, StreamingContext context) : base(serializationInfo, context)
        {}
    }
}