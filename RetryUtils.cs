using System;

namespace ImageConnect.Test.Func.Shared.Core.Utils
{
    /// <summary>
    /// The retry utils.
    /// </summary>
    public static class RetryUtils
    {
        /// <summary>
        /// Retries action if an exception thrown.
        /// </summary>
        /// <typeparam name="TException">The exception.</typeparam>
        /// <typeparam name="TRet">Result.</typeparam>
        /// <param name="action">Action.</param>
        /// <param name="numberOfTries">Number of tries.</param>
        /// <param name="actionAfterThrow">The action to be performed when an exception is thrown.</param>
        /// <returns>The result of action.</returns>
        public static TRet RetryIfThrown<TException, TRet>(Func<TRet> action, int numberOfTries, Action actionAfterThrow = null) where TException : Exception
        {
            TException lastException = null;

            for (var currentTry = 1; currentTry <= numberOfTries; currentTry++)
            {
                try
                {
                    return action();
                }
                catch (TException e)
                {
                    lastException = e;
                }
            }

            if (lastException != null)
            {
                actionAfterThrow?.Invoke();

                throw lastException;
            }

            return action();
        }

        /// <summary>
        /// Retries action that has no return type if an exception thrown.
        /// </summary>
        /// <typeparam name="TException">The exception.</typeparam>
        /// <param name="action">Action.</param>
        /// <param name="numberOfTries">Number of tries.</param>
        /// <param name="actionAfterThrow">The action to be performed when an exception is thrown.</param>
        public static void RetryIfThrown<TException>(Action action, int numberOfTries, Action actionAfterThrow = null) where TException : Exception
        {
            TException lastException = null;

            for (var currentTry = 1; currentTry <= numberOfTries; currentTry++)
            {
                try
                {
                    action();
                    return;
                }
                catch (TException e)
                {
                    lastException = e;
                }
            }

            if (lastException != null)
            {
                actionAfterThrow?.Invoke();

                throw lastException;
            }
        }
    }
}
