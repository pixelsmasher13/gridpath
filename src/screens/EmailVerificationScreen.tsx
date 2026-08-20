import React, { useState, useEffect } from 'react';
import {
  Box,
  VStack,
  Heading,
  Text,
  Button,
  Alert,
  AlertIcon,
  useColorModeValue,
  Flex,
  Image,
  Spinner,
  HStack,
  useToast,
  Link,
} from '@chakra-ui/react';
import { CheckCircleIcon, EmailIcon } from '@chakra-ui/icons';
import { useAuth } from '../Providers/AuthProvider';
import { sendVerificationEmail, reloadUser, logout } from '../services/auth';
import logoBlack from '@heelix-app/design/logo/logo-black.png';

export const EmailVerificationScreen: React.FC = () => {
  const [isChecking, setIsChecking] = useState(false);
  const [isResending, setIsResending] = useState(false);
  const [emailSent, setEmailSent] = useState(false);
  const [checkCount, setCheckCount] = useState(0);
  const { firebaseUser } = useAuth();
  const toast = useToast();

  const bgColor = useColorModeValue('gray.50', 'gray.900');
  const cardBg = useColorModeValue('white', 'gray.800');

  useEffect(() => {
    // Auto-check every 5 seconds, up to 12 times (1 minute)
    if (checkCount > 0 && checkCount < 12) {
      const timer = setTimeout(() => {
        handleCheckVerification(true);
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [checkCount]);

  const handleCheckVerification = async (isAutoCheck = false) => {
    if (!isAutoCheck) {
      setIsChecking(true);
    }

    try {
      const user = await reloadUser();

      if (user?.emailVerified) {
        toast({
          title: 'Email Verified!',
          description: 'Your email has been successfully verified. Logging you in...',
          status: 'success',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });

        // Reload the page to update auth state
        setTimeout(() => {
          window.location.reload();
        }, 1500);
      } else if (!isAutoCheck) {
        toast({
          title: 'Not Verified Yet',
          description: 'Please check your email and click the verification link',
          status: 'warning',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      }

      if (isAutoCheck) {
        setCheckCount(prev => prev + 1);
      }
    } catch (error) {
      if (!isAutoCheck) {
        toast({
          title: 'Error',
          description: 'Failed to check verification status',
          status: 'error',
          duration: 3000,
          isClosable: true,
          position: 'top',
        });
      }
    } finally {
      setIsChecking(false);
    }
  };

  const handleResendEmail = async () => {
    setIsResending(true);
    try {
      await sendVerificationEmail();
      setEmailSent(true);
      setCheckCount(1); // Start auto-checking

      toast({
        title: 'Verification Email Sent',
        description: `Check your inbox at ${firebaseUser?.email}`,
        status: 'success',
        duration: 5000,
        isClosable: true,
        position: 'top',
      });

      // Reset the "email sent" state after 30 seconds
      setTimeout(() => setEmailSent(false), 30000);
    } catch (error: any) {
      let errorMessage = 'Please try again later';

      if (error.code === 'auth/too-many-requests') {
        errorMessage = 'Too many requests. Please wait a few minutes before trying again.';
      }

      toast({
        title: 'Failed to Send Email',
        description: errorMessage,
        status: 'error',
        duration: 3000,
        isClosable: true,
        position: 'top',
      });
    } finally {
      setIsResending(false);
    }
  };

  const handleSignOut = async () => {
    try {
      await logout();
      window.location.reload();
    } catch (error) {
      console.error('Failed to sign out:', error);
    }
  };

  const handleStartAutoCheck = () => {
    setCheckCount(1);
    toast({
      title: 'Auto-checking started',
      description: 'We\'ll check your verification status automatically',
      status: 'info',
      duration: 3000,
      isClosable: true,
      position: 'top',
    });
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
        maxW="500px"
      >
        <VStack spacing={6} align="stretch">
          <Box textAlign="center">
            <Image src={logoBlack} alt="GridPath" w="50px" h="50px" mx="auto" mb={3} />
            <Heading size="lg" mb={2}>Verify Your Email</Heading>
            <Text color="gray.500" fontSize="sm">
              We've sent a verification link to
            </Text>
            <Text fontWeight="bold" fontSize="md" mt={1}>
              {firebaseUser?.email}
            </Text>
          </Box>

          <Alert status="info" borderRadius="md">
            <AlertIcon />
            <Box>
              <Text fontSize="sm">
                Please check your email and click the verification link to continue.
                The link will expire in 1 hour.
              </Text>
            </Box>
          </Alert>

          <VStack spacing={4}>
            {/* Status indicator */}
            <HStack justify="center" spacing={3}>
              {checkCount > 0 && checkCount < 12 ? (
                <>
                  <Spinner size="sm" color="blue.500" />
                  <Text fontSize="sm" color="gray.600">
                    Checking verification status...
                  </Text>
                </>
              ) : (
                <Text fontSize="sm" color="gray.600">
                  Click "I've Verified" after clicking the link in your email
                </Text>
              )}
            </HStack>

            {/* Primary action button */}
            <Button
              colorScheme="blue"
              size="lg"
              width="full"
              onClick={() => handleCheckVerification(false)}
              isLoading={isChecking}
              loadingText="Checking..."
              leftIcon={<CheckCircleIcon />}
            >
              I've Verified My Email
            </Button>

            {/* Secondary actions */}
            <VStack spacing={2} w="full">
              {!emailSent ? (
                <Button
                  variant="outline"
                  size="md"
                  width="full"
                  onClick={handleResendEmail}
                  isLoading={isResending}
                  loadingText="Sending..."
                  leftIcon={<EmailIcon />}
                >
                  Resend Verification Email
                </Button>
              ) : (
                <Button
                  variant="outline"
                  size="md"
                  width="full"
                  isDisabled
                  leftIcon={<CheckCircleIcon />}
                >
                  Email Sent ✓
                </Button>
              )}

              {checkCount === 0 && (
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={handleStartAutoCheck}
                >
                  Start auto-check
                </Button>
              )}
            </VStack>
          </VStack>

          {/* Help text */}
          <VStack spacing={2} pt={4} borderTop="1px" borderColor="gray.200">
            <Text fontSize="xs" color="gray.500">
              Didn't receive the email? Check your spam folder.
            </Text>
            <Text fontSize="xs" color="gray.500">
              Wrong email?{' '}
              <Link color="blue.500" onClick={handleSignOut}>
                Sign out
              </Link>{' '}
              and create a new account.
            </Text>
          </VStack>
        </VStack>
      </Box>
    </Flex>
  );
};