import { randomUUID } from "node:crypto";

export interface StoredArtifact {
    id: string;
    fileName: string;
    mimeType: string;
    buffer: Buffer;
    size: number;
    createdAt: Date;
    expiresAt: Date;
}

class ArtifactStore {

    private artifacts = new Map<string, StoredArtifact>();

    private readonly ttlMs =
        15 * 60 * 1000; // 15 dakika

    create(
        fileName: string,
        mimeType: string,
        buffer: Buffer
    ): StoredArtifact {

        const id = randomUUID();

        const now = new Date();

        const artifact: StoredArtifact = {
            id,
            fileName,
            mimeType,
            buffer,
            size: buffer.length,
            createdAt: now,
            expiresAt: new Date(
                now.getTime() + this.ttlMs
            )
        };

        this.artifacts.set(
            id,
            artifact
        );

        return artifact;
    }

    get(id: string): StoredArtifact | undefined {

        const artifact =
            this.artifacts.get(id);

        if (!artifact) {
            return undefined;
        }

        if (
            artifact.expiresAt.getTime()
            < Date.now()
        ) {

            this.artifacts.delete(id);

            return undefined;
        }

        return artifact;
    }

    delete(id: string): void {
        this.artifacts.delete(id);
    }

    cleanup(): void {

        const now = Date.now();

        for (
            const [id, artifact]
            of this.artifacts.entries()
        ) {

            if (
                artifact.expiresAt.getTime()
                < now
            ) {
                this.artifacts.delete(id);
            }
        }
    }
}

export const artifactStore =
    new ArtifactStore();