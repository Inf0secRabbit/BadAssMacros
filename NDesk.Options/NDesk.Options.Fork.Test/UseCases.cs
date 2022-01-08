using System;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace NDesk.Options.Fork.Test
{
    [TestFixture]
    public class UseCases
    {
        [TestCase("s|status", new[] { "-s", "-s" })]
        [TestCase("s|status", new[] { "-s", "/s" })]
        [TestCase("s|status", new[] { "-s", "/status" })]
        [TestCase("s|status", new[] { "-s", "--status" })]
        [TestCase("s|status", new[] { "/s", "/status" })]
        [TestCase("s|status", new[] { "/s", "--status" })]
        [TestCase("s|status", new[] { "/s", "/s" })]
        [TestCase("s|status", new[] { "--status", "/status" })]
        [TestCase("s|status", new[] { "--status", "--status" })]
        [TestCase("s|status", new[] { "/status", "/status" })]
        [TestCase("s|status", new[] { "/status", "--status", "-s", "/s" })]
        public void SameArgumentRecieveSameLambdaCalls(string optionCreate, string[] input)
        {
            // arrange
            var sut = new OptionSet();
            var action = Substitute.For<Action<string>>();

            sut.Add(optionCreate, action);

            // act
            var args2 = sut.Parse(input);

            // assert
            try
            {
                action.Received(input.Length).Invoke(Arg.Any<string>());
            }
            catch (Exception)
            {
                foreach (var s in input)
                {
                    Console.Write(s + " ");
                }

                throw;
            }
            foreach (var s in args2)
            {
                Console.Write(s + " ");
            }

        }

        [TestCase("s", new[] { "-f", "/status2", "--status2", "/s2", "unknown", "12", "s", "status" })]
        [TestCase("status", new[] { "-f", "/status2", "--status2", "/s2", "unknown", "12", "s", "status" })]
        [TestCase("s|status", new[] { "-f", "/status2", "--status2", "/s2", "unknown", "12", "s", "status" })]
        public void UnexpectedArguments(string optionCreate, string[] input)
        {
            // arrange
            var sut = new OptionSet();
            var action = Substitute.For<Action<string>>();

            sut.Add(optionCreate, action);

            // act
            var unexpectedArguments = sut.Parse(input);

            // assert
            //unexpectedArguments.ShouldAllBeEquivalentTo(input);
        }

        [TestCase("s", "-s")]
        [TestCase("s", "/s")]
        [TestCase("status", "--status")]
        [TestCase("status", "/status")]
        [TestCase("s|status", "-s")]
        [TestCase("s|status", "--status")]
        [TestCase("s|status", "/status")]
        public void OptionCallbackRecieveCall(string optionCreate, string optionUse)
        {
            // arrange
            var sut = new OptionSet();
            var input = new[] { optionUse };
            var action = Substitute.For<Action<string>>();

            sut.Add(optionCreate, action);

            // act
            sut.Parse(input);

            // assert
            action.Received(1).Invoke(Arg.Any<string>());
        }

        [TestCase("s=", "-s", "Ready")]
        [TestCase("s=", "/s", "Ready")]
        [TestCase("status=", "--status", "Ready")]
        [TestCase("status=", "/status", "Ready")]
        [TestCase("status=", "--status", "Ready")]
        [TestCase("s|status=", "-s", "Ready")]
        [TestCase("s|status=", "--status", "Ready")]
        public void OptionCallbackRecieveCallWithParameter(string optionCreate, string optionUse, string argument)
        {
            // arrange
            var sut = new OptionSet();
            var input = new[] { optionUse, argument };

            var action = Substitute.For<Action<string>>();
            var expectedActionArg = argument;

            sut.Add(optionCreate, action);

            // act
            sut.Parse(input);

            // assert
            action.Received(1)
                .Invoke(expectedActionArg);
        }
    }
}