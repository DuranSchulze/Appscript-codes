import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";

export default function TermsOfServicePage() {
    return (
        <div className="container mx-auto py-10 max-w-4xl px-4">
            <h1 className="text-4xl font-bold mb-8 text-center">Legal Information</h1>

            <div className="space-y-8">
                <section id="terms">
                    <Card>
                        <CardHeader>
                            <CardTitle className="text-2xl">Terms of Service</CardTitle>
                        </CardHeader>
                        <Separator className="mb-4" />
                        <CardContent className="space-y-4 text-muted-foreground">
                            <p className="text-sm">Last updated: {new Date().toLocaleDateString()}</p>
                            <p>
                                Welcome to Smart Transmittal. By accessing or using our service, you agree to be bound by these Terms of Service.
                            </p>

                            <div className="space-y-2">
                                <h3 className="text-lg font-semibold text-foreground">1. Acceptance of Terms</h3>
                                <p>
                                    By accessing and using this service, you accept and agree to be bound by the terms and provision of this agreement.
                                </p>
                            </div>

                            <div className="space-y-2">
                                <h3 className="text-lg font-semibold text-foreground">2. Use of Service</h3>
                                <p>
                                    You agree to use this service only for lawful purposes and in accordance with the stated purpose of document transmittal management.
                                </p>
                            </div>

                            <div className="space-y-2">
                                <h3 className="text-lg font-semibold text-foreground">3. User Responsibilities</h3>
                                <p>
                                    You are responsible for maintaining the confidentiality of your account information and for all activities that occur under your account.
                                </p>
                            </div>
                        </CardContent>
                    </Card>
                </section>
            </div>

            <div className="mt-12 text-center text-sm text-muted-foreground">
                <p>&copy; {new Date().getFullYear()} Smart Transmittal. All rights reserved.</p>
            </div>
        </div>
    );
}
