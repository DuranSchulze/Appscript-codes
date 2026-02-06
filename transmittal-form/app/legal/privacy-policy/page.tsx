import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";

export default function PrivacyPolicyPage() {
    return (
        <div className="container mx-auto py-10 max-w-4xl px-4">
            <h1 className="text-4xl font-bold mb-8 text-center">Legal Information</h1>

            <div className="space-y-8">
                <section id="privacy">
                    <Card>
                        <CardHeader>
                            <CardTitle className="text-2xl">Privacy Policy</CardTitle>
                        </CardHeader>
                        <Separator className="mb-4" />
                        <CardContent className="space-y-4 text-muted-foreground">
                            <p className="text-sm">Last updated: {new Date().toLocaleDateString()}</p>
                            <p>
                                Your privacy is important to us. It is Smart Transmittal&apos;s policy to respect your privacy regarding any information we may collect from you across our website.
                            </p>

                            <div className="space-y-2">
                                <h3 className="text-lg font-semibold text-foreground">1. Information We Collect</h3>
                                <p>
                                    We only ask for personal information when we truly need it to provide a service to you. We collect it by fair and lawful means, with your knowledge and consent.
                                </p>
                            </div>

                            <div className="space-y-2">
                                <h3 className="text-lg font-semibold text-foreground">2. Usage of Information</h3>
                                <p>
                                    We strictly use your personal information to provide the services you have requested. We do not share your personally identifying information publicly or with third-parties, except when required to by law.
                                </p>
                            </div>

                            <div className="space-y-2">
                                <h3 className="text-lg font-semibold text-foreground">3. Data Retention</h3>
                                <p>
                                    We only retain collected information for as long as necessary to provide you with your requested service.
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
