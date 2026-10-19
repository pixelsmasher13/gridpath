import React, { useState, type FC } from 'react';
import {
  Box,
  Button,
  VStack,
  Heading,
  Text,
  Link,
  Alert,
  AlertIcon,
  Flex,
  Input,
  FormControl,
  FormErrorMessage,
  Divider,
  InputGroup,
  InputRightElement,
  IconButton,
  HStack,
  Icon,
  useToast,
} from '@chakra-ui/react';
import { ViewIcon, ViewOffIcon } from '@chakra-ui/icons';
import { FcGoogle } from 'react-icons/fc';
import { useAuth } from '../Providers/AuthProvider';

type ViewMode = 'initial' | 'signup' | 'signin' | 'reset';

export const LoginScreen: FC = () => {
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState('');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [showConfirmPassword, setShowConfirmPassword] = useState(false);
  const [viewMode, setViewMode] = useState<ViewMode>('initial');
  const [resetEmailSent, setResetEmailSent] = useState(false);

  const { login, loginWithEmail, signUpWithEmail, resetPassword } = useAuth();
  const toast = useToast();

  // Login screen is locked to dark theme — matches the rest of the desktop
  // app (#1e1e1e bg, #252526 cards, #2a2a2a borders, #3363AD accent).
  // The previous useColorModeValue fallback rendered light/white, which
  // looked off-brand against the dark grid + chat panel.
  const bgColor = '#1e1e1e';
  const cardBg = '#252526';
  const borderColor = '#2a2a2a';
  const textColor = '#d4d4d4';
  const textSoft = '#9b9b9b';
  const textMute = '#7c7c7c';
  const accent = '#3363AD';
  const inputBg = '#161616';

  const validateEmail = (email: string) => {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  };

  const handleGoogleSignIn = async () => {
    setError('');
    setIsLoading(true);

    try {
      await login();
    } catch (err) {
      setError('Failed to sign in with Google. Please try again.');
    } finally {
      setIsLoading(false);
    }
  };

  const handleEmailContinue = (e: React.FormEvent) => {
    e.preventDefault();
    setError(''); // Clear any previous errors

    if (!validateEmail(email)) {
      setError('Please enter a valid email address');
      return;
    }

    // Clear error before moving to sign up view
    setError('');
    setPassword('');
    setConfirmPassword('');
    // Move to sign up view
    setViewMode('signup');
  };

  const handleEmailSignIn = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');

    if (!password || password.length < 6) {
      setError('Password must be at least 6 characters');
      return;
    }

    setIsLoading(true);
    try {
      await loginWithEmail(email, password);
    } catch (err: any) {
      console.error('Email sign in error:', err.code, err.message);

      let errorMessage = '';
      if (err.code === 'auth/user-not-found') {
        errorMessage = 'No account found with this email. Please sign up first.';
        toast({
          title: 'Account Not Found',
          description: errorMessage,
          status: 'error',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      } else if (err.code === 'auth/wrong-password') {
        errorMessage = 'Incorrect password. Please try again.';
        toast({
          title: 'Incorrect Password',
          description: errorMessage,
          status: 'error',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      } else if (err.code === 'auth/invalid-credential') {
        errorMessage = 'Invalid email or password. Please check and try again.';
        toast({
          title: 'Invalid Credentials',
          description: errorMessage,
          status: 'error',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      } else if (err.code === 'auth/invalid-email') {
        errorMessage = 'Invalid email address format.';
        toast({
          title: 'Invalid Email',
          description: errorMessage,
          status: 'error',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      } else if (err.code === 'auth/user-disabled') {
        errorMessage = 'This account has been disabled. Please contact support.';
        toast({
          title: 'Account Disabled',
          description: errorMessage,
          status: 'error',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      } else if (err.code === 'auth/too-many-requests') {
        errorMessage = 'Too many failed attempts. Please try again later.';
        toast({
          title: 'Too Many Attempts',
          description: errorMessage,
          status: 'error',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      } else {
        errorMessage = err.message || 'Failed to sign in. Please try again.';
        toast({
          title: 'Sign In Failed',
          description: errorMessage,
          status: 'error',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      }

      setError(errorMessage);
    } finally {
      setIsLoading(false);
    }
  };

  const handleEmailSignUp = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');

    if (!password || password.length < 6) {
      setError('Password must be at least 6 characters');
      return;
    }

    if (password !== confirmPassword) {
      setError('Passwords do not match');
      return;
    }

    setIsLoading(true);
    try {
      await signUpWithEmail(email, password);
      // If successful, Firebase auth state change will handle the redirect
    } catch (err: any) {
      console.error('Email sign up error:', err.code, err.message);

      // Always set loading to false when there's an error
      setIsLoading(false);

      let errorMessage = '';
      if (err.code === 'auth/email-already-in-use') {
        errorMessage = 'This email is already registered. Please sign in with your existing account.';
        // Show toast notification
        toast({
          title: 'Email Already Registered',
          description: 'This email is already registered. Please sign in instead.',
          status: 'error',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      } else if (err.code === 'auth/weak-password') {
        errorMessage = 'Password is too weak. Please use at least 6 characters.';
        toast({
          title: 'Weak Password',
          description: errorMessage,
          status: 'error',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      } else if (err.code === 'auth/invalid-email') {
        errorMessage = 'Invalid email address format.';
        toast({
          title: 'Invalid Email',
          description: errorMessage,
          status: 'error',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      } else if (err.code === 'auth/operation-not-allowed') {
        errorMessage = 'Email/password accounts are not enabled. Please contact support.';
        toast({
          title: 'Authentication Error',
          description: errorMessage,
          status: 'error',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      } else {
        errorMessage = err.message || 'Failed to create account. Please try again.';
        toast({
          title: 'Sign Up Failed',
          description: errorMessage,
          status: 'error',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      }

      // Also set the error in state for form display
      setError(errorMessage);
    }
  };

  const handlePasswordReset = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');

    if (!validateEmail(email)) {
      setError('Please enter a valid email address');
      return;
    }

    setIsLoading(true);
    try {
      await resetPassword(email);
      setResetEmailSent(true);
      setError('');
    } catch (err: any) {
      if (err.code === 'auth/user-not-found') {
        setError('No account found with this email address.');
      } else {
        setError(err.message || 'Failed to send reset email. Please try again.');
      }
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <Flex
      minH="100vh"
      align="center"
      justify="center"
      bg={bgColor}
    >
      <Box
        bg={cardBg}
        p={8}
        borderRadius="xl"
        boxShadow="xl"
        w="full"
        maxW="400px"
      >
        <VStack spacing={6} align="stretch">
          <Box textAlign="center">
            <Heading size="lg" mb={2} color={textColor}>Welcome to GridPath</Heading>
            <Text color={textSoft} fontSize="sm">
              {viewMode === 'initial' && 'Sign in to continue'}
              {viewMode === 'signup' && 'Create your account'}
              {viewMode === 'signin' && 'Sign in to your account'}
              {viewMode === 'reset' && 'Reset your password'}
            </Text>
          </Box>

          {error && (
            <Alert status="error" borderRadius="md" fontSize="sm">
              <AlertIcon />
              {error}
            </Alert>
          )}

          {resetEmailSent && (
            <Alert status="success" borderRadius="md" fontSize="sm">
              <AlertIcon />
              Password reset email sent! Check your inbox.
            </Alert>
          )}

          {/* Initial view with Google and email options */}
          {viewMode === 'initial' && (
            <VStack spacing={4}>
              <Button
                onClick={handleGoogleSignIn}
                size="md"
                width="full"
                variant="outline"
                borderColor={borderColor}
                color={textColor}
                bg="transparent"
                _hover={{ bg: '#2a2a2a' }}
                isLoading={isLoading}
                loadingText="Signing in..."
                leftIcon={<Icon as={FcGoogle} boxSize={5} />}
                fontWeight="medium"
                py={5}
              >
                Continue with Google
              </Button>

              <HStack w="full">
                <Divider borderColor={borderColor} />
                <Text fontSize="xs" color={textMute} px={2}>OR</Text>
                <Divider borderColor={borderColor} />
              </HStack>

              <form onSubmit={handleEmailContinue} style={{ width: '100%' }}>
                <VStack spacing={4}>
                  <FormControl isInvalid={!!error}>
                    <Input
                      type="email"
                      placeholder="Enter your email address"
                      value={email}
                      onChange={(e) => setEmail(e.target.value)}
                      isDisabled={isLoading}
                      size="md"
                      borderColor={borderColor}
                      bg={inputBg}
                      color={textColor}
                      _placeholder={{ color: '#555' }}
                      _focus={{ borderColor: accent, boxShadow: `0 0 0 1px ${accent}` }}
                    />
                    <FormErrorMessage fontSize="xs">{error}</FormErrorMessage>
                  </FormControl>

                  <Button
                    type="submit"
                    bg={accent}
                    color="#fff"
                    _hover={{ bg: '#4275c4' }}
                    size="md"
                    width="full"
                    isLoading={isLoading}
                    fontWeight="medium"
                  >
                    Continue
                  </Button>
                </VStack>
              </form>
            </VStack>
          )}

          {/* Sign up view */}
          {viewMode === 'signup' && (
            <form onSubmit={handleEmailSignUp} style={{ width: '100%' }}>
              <VStack spacing={4}>
                <Box w="full">
                  <Text fontSize="sm" fontWeight="medium" mb={2} color={textColor}>
                    Email
                  </Text>
                  <Text fontSize="sm" color={textSoft} mb={3}>
                    {email}
                  </Text>
                </Box>

                <VStack spacing={3} w="full">
                  <Text fontSize="sm" color={textSoft} alignSelf="start">
                    Set your password to continue
                  </Text>

                  <FormControl isInvalid={!!error && error.includes('password')}>
                    <InputGroup size="sm">
                      <Input
                        type={showPassword ? 'text' : 'password'}
                        placeholder="Password (min 6 characters)"
                        value={password}
                        onChange={(e) => setPassword(e.target.value)}
                        isDisabled={isLoading}
                        fontSize="sm"
                        bg={inputBg}
                        color={textColor}
                        borderColor={borderColor}
                        _placeholder={{ color: '#555' }}
                        _focus={{ borderColor: accent, boxShadow: `0 0 0 1px ${accent}` }}
                      />
                      <InputRightElement>
                        <IconButton
                          aria-label={showPassword ? 'Hide password' : 'Show password'}
                          icon={showPassword ? <ViewOffIcon /> : <ViewIcon />}
                          onClick={() => setShowPassword(!showPassword)}
                          variant="ghost"
                          size="xs"
                          color={textSoft}
                        />
                      </InputRightElement>
                    </InputGroup>
                  </FormControl>

                  <FormControl isInvalid={!!error && error.includes('match')}>
                    <InputGroup size="sm">
                      <Input
                        type={showConfirmPassword ? 'text' : 'password'}
                        placeholder="Confirm password"
                        value={confirmPassword}
                        onChange={(e) => setConfirmPassword(e.target.value)}
                        isDisabled={isLoading}
                        fontSize="sm"
                        bg={inputBg}
                        color={textColor}
                        borderColor={borderColor}
                        _placeholder={{ color: '#555' }}
                        _focus={{ borderColor: accent, boxShadow: `0 0 0 1px ${accent}` }}
                      />
                      <InputRightElement>
                        <IconButton
                          aria-label={showConfirmPassword ? 'Hide password' : 'Show password'}
                          icon={showConfirmPassword ? <ViewOffIcon /> : <ViewIcon />}
                          onClick={() => setShowConfirmPassword(!showConfirmPassword)}
                          variant="ghost"
                          size="xs"
                          color={textSoft}
                        />
                      </InputRightElement>
                    </InputGroup>
                  </FormControl>
                </VStack>

                <Button
                  type="submit"
                  bg={accent}
                  color="#fff"
                  _hover={{ bg: '#4275c4' }}
                  size="md"
                  width="full"
                  isLoading={isLoading}
                  loadingText="Creating account..."
                  fontWeight="medium"
                >
                  Create Account
                </Button>

                <Text fontSize="sm" color={textSoft}>
                  Existing user?{' '}
                  <Link
                    color={accent}
                    fontWeight="medium"
                    onClick={() => {
                      setViewMode('signin');
                      setError('');
                      setPassword('');
                      setConfirmPassword('');
                    }}
                  >
                    Sign in here
                  </Link>
                </Text>

                {/* Show sign in button prominently if email already exists */}
                {error && (error.includes('already registered') || error.includes('already exists')) && (
                  <Button
                    variant="solid"
                    bg={accent}
                    color="#fff"
                    _hover={{ bg: '#4275c4' }}
                    size="sm"
                    width="full"
                    onClick={() => {
                      setViewMode('signin');
                      setError('');
                      setPassword('');
                      setConfirmPassword('');
                    }}
                  >
                    Go to Sign In
                  </Button>
                )}
              </VStack>
            </form>
          )}

          {/* Sign in view */}
          {viewMode === 'signin' && (
            <form onSubmit={handleEmailSignIn} style={{ width: '100%' }}>
              <VStack spacing={4}>
                <FormControl>
                  <Input
                    type="email"
                    placeholder="Email"
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
                    isDisabled={isLoading}
                    size="sm"
                    fontSize="sm"
                    bg={inputBg}
                    color={textColor}
                    borderColor={borderColor}
                    _placeholder={{ color: '#555' }}
                    _focus={{ borderColor: accent, boxShadow: `0 0 0 1px ${accent}` }}
                  />
                </FormControl>

                <FormControl isInvalid={!!error && error.includes('password')}>
                  <InputGroup size="sm">
                    <Input
                      type={showPassword ? 'text' : 'password'}
                      placeholder="Password"
                      value={password}
                      onChange={(e) => setPassword(e.target.value)}
                      isDisabled={isLoading}
                      fontSize="sm"
                      bg={inputBg}
                      color={textColor}
                      borderColor={borderColor}
                      _placeholder={{ color: '#555' }}
                      _focus={{ borderColor: accent, boxShadow: `0 0 0 1px ${accent}` }}
                    />
                    <InputRightElement>
                      <IconButton
                        aria-label={showPassword ? 'Hide password' : 'Show password'}
                        icon={showPassword ? <ViewOffIcon /> : <ViewIcon />}
                        onClick={() => setShowPassword(!showPassword)}
                        variant="ghost"
                        size="xs"
                        color={textSoft}
                      />
                    </InputRightElement>
                  </InputGroup>
                </FormControl>

                <Button
                  type="submit"
                  bg={accent}
                  color="#fff"
                  _hover={{ bg: '#4275c4' }}
                  size="md"
                  width="full"
                  isLoading={isLoading}
                  loadingText="Signing in..."
                  fontWeight="medium"
                >
                  Sign In
                </Button>

                <VStack spacing={2} w="full">
                  <Link
                    fontSize="sm"
                    color={accent}
                    onClick={() => {
                      setViewMode('reset');
                      setError('');
                    }}
                  >
                    Forgot password?
                  </Link>

                  <Text fontSize="sm" color={textSoft}>
                    New user?{' '}
                    <Link
                      color={accent}
                      fontWeight="medium"
                      onClick={() => {
                        // Go back to initial page with Google option
                        setViewMode('initial');
                        setError('');
                        setEmail('');
                        setPassword('');
                      }}
                    >
                      Create account
                    </Link>
                  </Text>

                  <Divider my={2} borderColor={borderColor} />

                  <Link
                    fontSize="xs"
                    color={textMute}
                    onClick={() => {
                      setViewMode('initial');
                      setError('');
                      setEmail('');
                      setPassword('');
                    }}
                  >
                    Or sign in with Google
                  </Link>
                </VStack>
              </VStack>
            </form>
          )}

          {/* Password reset view */}
          {viewMode === 'reset' && (
            <form onSubmit={handlePasswordReset} style={{ width: '100%' }}>
              <VStack spacing={4}>
                <Text fontSize="sm" color={textSoft}>
                  Enter your email to receive a password reset link
                </Text>

                <FormControl isInvalid={!!error}>
                  <Input
                    type="email"
                    placeholder="Email address"
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
                    isDisabled={isLoading}
                    size="sm"
                    fontSize="sm"
                    bg={inputBg}
                    color={textColor}
                    borderColor={borderColor}
                    _placeholder={{ color: '#555' }}
                    _focus={{ borderColor: accent, boxShadow: `0 0 0 1px ${accent}` }}
                  />
                </FormControl>

                <Button
                  type="submit"
                  bg={accent}
                  color="#fff"
                  _hover={{ bg: '#4275c4' }}
                  size="md"
                  width="full"
                  isLoading={isLoading}
                  loadingText="Sending..."
                  fontWeight="medium"
                >
                  Send Reset Email
                </Button>

                <Link
                  fontSize="sm"
                  color={accent}
                  onClick={() => {
                    setViewMode('signin');
                    setError('');
                    setResetEmailSent(false);
                  }}
                >
                  Back to sign in
                </Link>
              </VStack>
            </form>
          )}

          {/* Footer links - only on initial view */}
          {viewMode === 'initial' && (
            <>
              <Divider borderColor={borderColor} />
              <VStack spacing={3}>
                <Text fontSize="xs" color={textMute} textAlign="center">
                  By continuing, you agree to our{' '}
                  <Link color={accent} href="https://gridpath.dev/terms" isExternal>
                    Terms
                  </Link>{' '}
                  and{' '}
                  <Link color={accent} href="https://gridpath.dev/privacy" isExternal>
                    Privacy Policy
                  </Link>
                </Text>

                <Text fontSize="xs" color={textMute}>
                  Need help?{' '}
                  <Link
                    color={accent}
                    href="https://gridpath.dev/contact"
                    isExternal
                  >
                    Contact support
                  </Link>
                </Text>
              </VStack>
            </>
          )}
        </VStack>
      </Box>
    </Flex>
  );
};